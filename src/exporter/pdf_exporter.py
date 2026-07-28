"""
Step 4: Export & PDF Conversion Pipeline Module.
Handles PPTX saving and automatic PDF conversion via LibreOffice CLI with PowerPoint COM fallback.
"""

import os
import shutil
import logging
import subprocess
from pathlib import Path
from typing import Tuple

logger = logging.getLogger("PDFExporter")


class PDFExporter:
    """Exports PPTX presentation and converts to PDF."""

    def __init__(self):
        self.libreoffice_cmd = self._find_libreoffice()

    def convert_pptx_to_pdf(self, pptx_path: str, pdf_output_path: str = None) -> Tuple[bool, str, str]:
        """
        Converts a PPTX file to PDF using available converters.
        Returns: (success_boolean, pdf_file_path, conversion_method_used)
        """
        abs_pptx = os.path.abspath(pptx_path)
        if not os.path.exists(abs_pptx):
            raise FileNotFoundError(f"PPTX file not found for PDF conversion: {abs_pptx}")

        if not pdf_output_path:
            pdf_output_path = str(Path(abs_pptx).with_suffix(".pdf"))
        abs_pdf = os.path.abspath(pdf_output_path)

        logger.info(f"Starting Step 4: Converting '{abs_pptx}' to PDF...")

        # Method 1: Primary - LibreOffice CLI
        if self.libreoffice_cmd:
            success, err = self._convert_via_libreoffice(abs_pptx, abs_pdf)
            if success:
                logger.info(f"PDF conversion successful via LibreOffice CLI -> '{abs_pdf}'")
                return True, abs_pdf, "LibreOffice CLI"
            else:
                logger.warning(f"LibreOffice CLI conversion attempt failed: {err}")

        # Method 2: Windows Native Fallback - PowerPoint COM automation
        if os.name == "nt":
            success, err = self._convert_via_powerpoint_com(abs_pptx, abs_pdf)
            if success:
                logger.info(f"PDF conversion successful via PowerPoint COM -> '{abs_pdf}'")
                return True, abs_pdf, "PowerPoint COM (Windows Native)"
            else:
                logger.warning(f"PowerPoint COM conversion attempt failed: {err}")

        # Method 3: If no converter succeeded
        failure_msg = (
            "PDF conversion could not be completed automatically. "
            "Attempted methods: [LibreOffice CLI, PowerPoint COM]. "
            "Ensure LibreOffice is installed or MS PowerPoint is active."
        )
        logger.error(failure_msg)
        return False, "", failure_msg

    def _convert_via_libreoffice(self, abs_pptx: str, abs_pdf: str) -> Tuple[bool, str]:
        """Runs LibreOffice in headless mode to convert PPTX to PDF."""
        out_dir = os.path.dirname(abs_pdf)
        cmd = [self.libreoffice_cmd, "--headless", "--convert-to", "pdf", abs_pptx, "--outdir", out_dir]
        try:
            res = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
            if res.returncode == 0:
                expected_out = str(Path(abs_pptx).with_suffix(".pdf"))
                if os.path.exists(expected_out) and expected_out != abs_pdf:
                    shutil.move(expected_out, abs_pdf)
                return os.path.exists(abs_pdf), ""
            return False, res.stderr or res.stdout
        except Exception as e:
            return False, str(e)

    def _convert_via_powerpoint_com(self, abs_pptx: str, abs_pdf: str) -> Tuple[bool, str]:
        """Converts PPTX to PDF using win32com PowerPoint application on Windows."""
        try:
            import win32com.client
            import pythoncom
            pythoncom.CoInitialize()
            
            ppt_app = win32com.client.Dispatch("PowerPoint.Application")
            # 32 = ppSaveAsPDF constant in PowerPoint COM API
            presentation = ppt_app.Presentations.Open(abs_pptx, False, False, False)
            presentation.SaveCopyAs(abs_pdf, 32)
            presentation.Close()
            ppt_app.Quit()
            pythoncom.CoUninitialize()
            return os.path.exists(abs_pdf), ""
        except Exception as e:
            return False, str(e)

    @staticmethod
    def _find_libreoffice() -> str:
        """Finds LibreOffice soffice executable in system PATH or standard installation locations."""
        soffice = shutil.which("soffice") or shutil.which("libreoffice")
        if soffice:
            return soffice

        # Standard Windows installation paths
        win_paths = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
        for p in win_paths:
            if os.path.exists(p):
                return p
        return None
