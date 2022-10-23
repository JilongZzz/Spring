import shutil
import os

OUT_DIR = 'work'

if os.path.exists(OUT_DIR):
    shutil.rmtree(OUT_DIR)
os.mkdir(OUT_DIR)

pwd = os.getcwd()
script = os.path.join(pwd, "rename", 'rename.bat')
out_script = os.path.join(pwd, OUT_DIR, 'step1_rename.bat')
print(script)
print(out_script)
shutil.copy(script, out_script)

script = os.path.join(pwd, "rank", 'rank.py')
out_script = os.path.join(pwd, OUT_DIR, 'step2_rank.py')
shutil.copy(script, out_script)

script = os.path.join(pwd, "rank", 'null.csv')
out_script = os.path.join(pwd, OUT_DIR, 'null.csv')
shutil.copy(script, out_script)

script = os.path.join(pwd, "splite_xls_sheet", 'splite_xls_sheet.py')
out_script = os.path.join(pwd, OUT_DIR, 'step3_splite_xls_sheet.py')
shutil.copy(script, out_script)
