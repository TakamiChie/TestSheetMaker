import unittest

import yaml

import src.main as main

class TestRearrangeCells(unittest.TestCase):

  TEST_CELLS = [
    ["No", "l", "m", "s", "cond", "proc", "exp"] + ["tester", "checker", "date", "result"] * 2,
    ["1-1-1", "test", "test", "test", "testcond", "testproc", "testexp", "1", "2", "3", "4", "5", "6", "7", "8"],
    ["1-1-2", "test", "test", "test2", "testcond", "testproc", "testexp", "1", "2", "3", "4", "5", "6", "7", "8"],
    ["1-2-1", "test", "test3", "test", "testcond", "testproc", "testexp", "1", "2", "3", "4", "5", "6", "7", "8"],
  ]
  def setUp(self) -> None:
    with open("./tests/test_config.yml") as f: self.config = yaml.safe_load(f)
    return super().setUp()

  def test_donothing(self):
    data = main.rearrange_cells(self.config["Headers"], self.TEST_CELLS, ["no", "itemname", "content", "results"])
    self.assertEqual(len(data), len(self.TEST_CELLS))
    for i, d in enumerate(data):
      self.assertEqual(d, self.TEST_CELLS[i], msg=f"assert Index {i}")

  def test_reverse(self):
    data = main.rearrange_cells(self.config["Headers"], self.TEST_CELLS, ["results", "content", "itemname", "no"])
    self.assertEqual(len(data), len(self.TEST_CELLS))
    for i, d in enumerate(data):
      self.assertEqual(d, self.TEST_CELLS[i][7:19] + self.TEST_CELLS[i][4:7] + self.TEST_CELLS[i][1:4] + self.TEST_CELLS[i][0:1], 
        msg=f"assert Index {i}")

  def test_working(self):
    data = main.rearrange_cells(self.config["Headers"], self.TEST_CELLS, ["itemname", "no", "results", "content"])
    self.assertEqual(len(data), len(self.TEST_CELLS))
    for i, d in enumerate(data):
      self.assertEqual(d, self.TEST_CELLS[i][1:4] + self.TEST_CELLS[i][0:1] + self.TEST_CELLS[i][7:19] + self.TEST_CELLS[i][4:7],
        msg=f"assert Index {i}")

  def test_error(self):
    with self.assertRaises(Exception):
      data = main.rearrange_cells(self.config["Headers"], self.TEST_CELLS, ["no", "itemname", "content", "result"])
