#!/bin/bash
# Double-click this file to open the PDF Tool.
# Uses absolute paths, so this file works even if you move it to the Desktop.
cd "~/Developer/claude code projects/pdf-tool" || exit 1
exec "~/Developer/claude code projects/pdf-tool/.xlsxenv/bin/python" "~/Developer/claude code projects/pdf-tool/pdftool.py"
