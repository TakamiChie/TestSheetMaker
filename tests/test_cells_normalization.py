import unittest
import src.main as main

class TestCellsNormalization(unittest.TestCase):
  def test_simple(self):
    testdata = [{
      "items": ["test", "testb", "testc"],
      "exams": {"aaa": "bbb", "bbb": "ccc"}
    }, {
      "items": ["test", "testb", "testd"],
      "exams": {"aaa": "bbb", "bbb": "ccc"}
    }]
    headers = ["l", "m", "s"]
    data = main.cells_normalization(headers, testdata)
    self.assertEquals(len(data), 3)
    self.assertEquals(data[0], ["No", "l", "m", "s", "aaa", "bbb"])
    self.assertEquals(data[1], ["1-1-1", "test", "testb", "testc", "bbb", "ccc"])
    self.assertEquals(data[2], ["1-1-2", "test", "testb", "testd", "bbb", "ccc"])

  def test_n_enough_items(self):
    testdata = [{
      "items": ["test", "testb", "testc"],
      "exams": {"aaa": "bbb", "bbb": "ccc"}
    },{
      "items": ["test", "testd"],
      "exams": {"aaa": "bbb", "ddd": "aaa", "bbb": "ccc"}
    },
    ]
    headers = ["l", "m", "s"]
    data = main.cells_normalization(headers, testdata)
    self.assertEquals(len(data), 3)
    self.assertEquals(data[0], ["No", "l", "m", "s", "aaa", "bbb", "ddd"])
    self.assertEquals(data[1], ["1-1-1", "test", "testb", "testc", "bbb", "ccc"])
    self.assertEquals(data[2], ["1-2-1", "test", "testd", "", "bbb", "ccc", "aaa"])

  def test_error_headers(self):
    testdata = [{
      "items": ["test", "testb", "testc", "testd"],
      "exams": {"aaa": "bbb", "bbb": "ccc"}
    }]
    headers = ["l", "m", "s"]
    with self.assertRaises(Exception):
      main.cells_normalization(headers, testdata)
