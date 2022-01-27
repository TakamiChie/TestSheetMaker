import unittest
import src.main as main

class TestAddExamcells(unittest.TestCase):
  CONFIG_DATA = {
    "PrintCount": 2,
    "Labels": ["tester","checker","date", "result"]
  }
  TEST_CELLS = [
    ["a", "b", "c", "d"],
    ["e", "f", "g", "h"],
    ["i", "j", "k", "l"],
  ]

  def test_right(self):
    data = main.add_examcells(self.CONFIG_DATA, self.TEST_CELLS)
    self.assertEqual(len(data[0]), 12)
    self.assertEqual(data[0], ["a", "b", "c", "d"] + (["tester", "checker", "date", "result"] * 2))
    self.assertEqual(data[1], ["e", "f", "g", "h"] + ([""] * 8))
    self.assertEqual(data[2], ["i", "j", "k", "l"] + ([""] * 8))
