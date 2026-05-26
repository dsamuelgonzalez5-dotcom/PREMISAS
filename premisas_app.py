import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta, date
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO
import re
import copy
import traceback
import sys
import gc

st.set_page_config(page_title="Test", page_icon="⚡", layout="wide")
st.write("## ✅ Todos los imports OK")
st.write(f"pandas {pd.__version__} | numpy {np.__version__} | openpyxl {openpyxl.__version__}")
