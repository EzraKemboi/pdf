from flask import Flask, request, jsonify, send_file
from PyPDF2 import PdfReader, PdfWriter
from docx import Document
import comtypes.client