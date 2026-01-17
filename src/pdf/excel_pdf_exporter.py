import logging
import os
import time

logger = logging.getLogger(__name__)

try:
    import pythoncom
    import win32com.client
    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False
    logger.warning(
        "win32com/pythoncom not available. "
        "PDF export requires Windows + Excel + pywin32."
    )


class ExcelPdfExporter:
    """
    Context-managed Excel → PDF exporter (Windows only).

    Key behavior:
    - Opens workbook in real Excel
    - Forces full recalculation (dependency rebuild)
    - Waits for calculation to complete
    - Optionally saves workbook
    - Exports either a worksheet or whole workbook to PDF
    - Cleans up Excel COM properly
    """

    def __init__(
        self,
        visible: bool = False,
        timeout_seconds: int = 30,
        force_full_recalc: bool = True,
        save_before_export: bool = True,
        use_fresh_instance: bool = True,
    ):
        self.visible = visible
        self.timeout_seconds = int(timeout_seconds)
        self.force_full_recalc = bool(force_full_recalc)
        self.save_before_export = bool(save_before_export)
        self.use_fresh_instance = bool(use_fresh_instance)

        self.excel = None
        self.logger = logging.getLogger(__name__)

    def __enter__(self):
        if not WIN32COM_AVAILABLE:
            raise RuntimeError(
                "ExcelPdfExporter requires Windows + Microsoft Excel + pywin32.\n"
                "Install pywin32:\n"
                "  pip install pywin32\n"
                "Then re-run on a Windows machine with Excel installed."
            )

        self.logger.info("Starting Excel COM application")
        pythoncom.CoInitialize()

        # DispatchEx => isolated instance (less interference with user-open Excel)
        if self.use_fresh_instance:
            self.excel = win32com.client.DispatchEx("Excel.Application")
        else:
            self.excel = win32com.client.Dispatch("Excel.Application")

        self.excel.Visible = self.visible
        self.excel.DisplayAlerts = False
        self.excel.AskToUpdateLinks = False
        self.excel.EnableEvents = False

        # Ensure automatic calc is enabled; we still force a calc later.
        # -4105 = xlCalculationAutomatic
        try:
            self.excel.Calculation = -4105
        except Exception:
            pass

        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        try:
            if self.excel:
                self.logger.info("Closing Excel COM application")
                try:
                    self.excel.Quit()
                except Exception:
                    pass
        finally:
            self.excel = None
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def export(
        self,
        workbook_path: str,
        output_pdf_path: str,
        sheet_name: str | None = None,
        quality: str = "standard",
        open_after_publish: bool = False,
        ignore_print_areas: bool = False,
    ) -> str:
        """
        Export workbook (or a single sheet) to PDF.

        Args:
            workbook_path: Path to .xlsx
            output_pdf_path: Path to output PDF
            sheet_name: If provided, exports that worksheet; otherwise exports whole workbook
            quality: "standard" or "minimum"
            open_after_publish: Whether to open PDF after export
            ignore_print_areas: If True, ignores print areas

        Returns:
            Absolute path to created PDF
        """
        if not self.excel:
            raise RuntimeError("ExcelPdfExporter not initialized. Use within a 'with' block.")

        workbook_path = os.path.abspath(workbook_path)
        output_pdf_path = os.path.abspath(output_pdf_path)

        if not os.path.exists(workbook_path):
            raise FileNotFoundError(f"Workbook not found: {workbook_path}")

        out_dir = os.path.dirname(output_pdf_path)
        if out_dir and not os.path.exists(out_dir):
            os.makedirs(out_dir, exist_ok=True)

        self.logger.info("Exporting PDF: %s -> %s", workbook_path, output_pdf_path)

        # Excel constants (avoid relying on constants generation)
        xlTypePDF = 0
        xlQualityStandard = 0
        xlQualityMinimum = 1
        quality_map = {"standard": xlQualityStandard, "minimum": xlQualityMinimum}

        wb = None
        try:
            wb = self.excel.Workbooks.Open(
                workbook_path,
                UpdateLinks=0,
                ReadOnly=False,
                IgnoreReadOnlyRecommended=True,
                AddToMru=False,
            )

            # Force calc BEFORE export (this addresses your “dropdown click then values appear” issue)
            self._force_recalc_and_wait()

            if self.save_before_export:
                try:
                    wb.Save()
                except Exception as e:
                    # Saving is useful, but not always required. Still, warn loudly if it fails.
                    self.logger.warning("Workbook save before export failed: %s", e)

            # Export either a specific worksheet or the full workbook
            if sheet_name:
                ws = wb.Worksheets(sheet_name)
                ws.ExportAsFixedFormat(
                    Type=xlTypePDF,
                    Filename=output_pdf_path,
                    Quality=quality_map.get(quality, xlQualityStandard),
                    IncludeDocProperties=True,
                    IgnorePrintAreas=bool(ignore_print_areas),
                    OpenAfterPublish=bool(open_after_publish),
                )
            else:
                wb.ExportAsFixedFormat(
                    Type=xlTypePDF,
                    Filename=output_pdf_path,
                    Quality=quality_map.get(quality, xlQualityStandard),
                    IncludeDocProperties=True,
                    IgnorePrintAreas=bool(ignore_print_areas),
                    OpenAfterPublish=bool(open_after_publish),
                )

            if not os.path.exists(output_pdf_path):
                raise RuntimeError(f"PDF export failed (no file created): {output_pdf_path}")

            self.logger.info("PDF exported successfully: %s", output_pdf_path)
            return output_pdf_path

        finally:
            if wb is not None:
                try:
                    wb.Close(SaveChanges=False)
                except Exception:
                    pass

    def _force_recalc_and_wait(self):
        """
        Force a full calculation and wait until Excel reports it's done.
        """
        if self.force_full_recalc:
            # Rebuild dependency tree + recalc
            self.excel.CalculateFullRebuild()
        else:
            self.excel.Calculate()

        start = time.time()
        # CalculationState: 0=Done, 1=Calculating, 2=Pending
        while getattr(self.excel, "CalculationState", 0) != 0:
            if time.time() - start > self.timeout_seconds:
                raise TimeoutError(
                    f"Excel calculation did not complete within {self.timeout_seconds}s"
                )
            time.sleep(0.2)
