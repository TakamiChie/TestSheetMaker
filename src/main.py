from argparse import ArgumentParser
from audioop import reverse
from ctypes import alignment
import re
from turtle import color
from typing import Any
from pathlib import Path

import openpyxl
import openpyxl.styles as styles
import yaml

def generate_testlist(lines: str) -> list[dict[list[str] | dict[str]]]:
  level = 0
  examsmap = []
  itemmap = []
  previoustest = {}
  currenttest = {}
  section = ""
  textbuf = []
  for line in lines.split("\n"):
    if m := re.match(r"^\s*(#+)\s*(.*)$", line):
      # change item
      if textbuf != [] and section != "":
        currenttest[section] = textbuf
      if itemmap != [] and currenttest != {}:
        examsmap.append({
          "items": itemmap.copy(),
          "exams": currenttest,
        })
      # new item
      ml = len(m[1])
      if level == ml:
        itemmap[-1] = m[2]
      elif level + 1 == ml:
        itemmap.append(m[2])
        level+=1
      elif level > ml:
        itemmap = itemmap[0:ml]
        itemmap[-1] = m[2]
        level=ml
      else:
        raise Exception("Incorrect test vote data.")
      section = ""
      if currenttest != {}:
        previoustest = currenttest.copy()
      currenttest = {}
      textbuf = []
    elif m := re.match(r"^\s*::\s*(.*?)\s*(&&)?$", line):
      # change section
      if textbuf != [] and section != "":
        currenttest[section] = textbuf
      # new section
      section = m[1]
      if m[2] == "&&" and section in previoustest:
        textbuf = previoustest[section]
      else:
        textbuf = []
    elif re.match(r"^\s*$", line):
      # ignore blank line.
      pass
    else:
      textbuf += [line.strip()]
  if textbuf != [] and section != "":
    currenttest[section] = textbuf
  if itemmap != [] and currenttest != {}:
    examsmap.append({
      "items": itemmap.copy(),
      "exams": currenttest,
    })
  return examsmap

def cells_normalization(testitemslabel: list[str], examsmap: list[dict[list[str] | dict[str]]]) -> list[list[str]]:
  tilcount = len(testitemslabel)
  lines = []
  header = {}
  ids = [0] * tilcount
  prevrowname = [""] * tilcount
  for exam in examsmap:
    line = []
    # itemname
    if tilcount > len(exam["items"]):
      line = exam["items"] + [""] * (tilcount - len(exam["items"]))
    elif tilcount < len(exam["items"]):
      raise Exception("Incorrect test data.") 
    else:
      line = exam["items"]
    # name count
    changed = False
    for i, n in enumerate(exam["items"]):
      if prevrowname[i] != n and not changed:
        ids[i] += 1
        if i < tilcount:
          ids[(i + 1):] = [1] * (tilcount - i - 1)
        changed = True
      prevrowname[i] = n
    line.insert(0, "-".join(map(str, ids)))
    # preload exams
    for ex in exam["exams"].items():
      if not ex[0] in header:
        header[ex[0]] = len(header)
    examdata = [""] * len(header)
    for n, v in exam["exams"].items():
      examdata[header[n]] = v
    lines.append(line + examdata)

  line = ["No"] + testitemslabel
  for h in header.keys():
    line.append(h)
  lines.insert(0, line)
  return lines
  
def create_excel(config:dict[Any], cells: list[list[str]], path: Path) -> None:
  wb = openpyxl.Workbook()
  ws = wb.worksheets[-1]
  ws.title = config["Sheet"]["Name"]
  # define styles
  headdesign = styles.PatternFill(patternType='solid', fgColor=config["Headers"]["BackColor"], bgColor=config["Headers"]["BackColor"])
  headfont = styles.Font(color=config["Headers"]["TextColor"])
  side = styles.Side(style="thin", color="000000")
  border = styles.Border(side, side, side, side)
  # fill header
  if "TestResult" in config["Headers"]:
    for c in range(config["Headers"]["TestResult"]["PrintCount"]):
      for l in config["Headers"]["TestResult"]["Labels"]:
        cells[0].append(f"{c+1}:{l}")
  # output cells
  for r, line in enumerate(cells):
    for c, cell in enumerate(line):
      cellobj = ws.cell(r + 1, c + 1)
      if type(cell) is list:
        cell = "\n".join(cell)
      cellobj.value = expandvars(cell, config["Consts"])
      align = {
        "vertical": "top",
        "horizontal": "justify"
      }
      if "\n" in cell:
        align["wrapText"] = True
      cellobj.alignment = styles.Alignment(**align) 
      if r == 0:
        # extension width
        calcsize = (len(cell) + 2) * 1.4
        dimensions = ws.column_dimensions[cellobj.column_letter]
        if not "\n" in cell and dimensions.width < calcsize:
           dimensions.width = calcsize
        # set style
        cellobj.fill = headdesign
        cellobj.font = headfont
      cellobj.border = border
    if len(cells[0]) - len(line) > 0:
      for c in range(len(cells[0]) - len(line)):
        ws.cell(r + 1, c + 1 + len(line)).border = border

  if not path.parent.exists(): path.parent.mkdir()
  wb.save(path)
  
def expandvars(text: str, consts: dict[str,str]):
  for n, v in consts.items():
    text = text.replace("{{" +n + "}}", v)
  return text

if __name__ == "__main__":
  p = ArgumentParser(description="Test Sheet Creation Tool")
  p.add_argument("tests", type=str, help="Markdown file that defines a test item.")
  p.add_argument("-o", "--out", required=True, type=str, help="Excel file output destination.")
  p.add_argument("-c", "--config", default="sample/config.yml", type=str, help="Configured file that defines basic information of the test vote.")
  args = p.parse_args()

  with open(args.config, mode="r", encoding="utf-8") as f: config = yaml.safe_load(f)
  with open(args.tests, mode="r", encoding="utf-8") as f: lines = f.read()

  exams = generate_testlist(lines)
  cells = cells_normalization(config["Headers"]["TestItemsLabel"], exams)
  create_excel(config, cells, path=Path(args.out))

