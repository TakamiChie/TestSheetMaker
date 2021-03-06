import unittest
import src.main as main

class TestGenerateTestlist(unittest.TestCase):
  def test_simple(self):
    testdata = """
    # test
    ## testb
    ### testc
    :: aaa
    bbb
    :: bbb
    ccc"""
    data = main.generate_testlist(testdata)
    self.assertEqual(len(data), 1)
    self.assertEqual(data[0]["items"], ["test", "testb", "testc"])
    self.assertEqual(data[0]["exams"], {"aaa": ["bbb"], "bbb": ["ccc"]})

  def test_simple_blankline(self):
    testdata = """
    # test
    ## testb
    ### testc
    :: aaa
    bbb

    :: bbb
    ccc
    """
    data = main.generate_testlist(testdata)
    self.assertEqual(len(data), 1)
    self.assertEqual(data[0]["items"], ["test", "testb", "testc"])
    self.assertEqual(data[0]["exams"], {"aaa": ["bbb"], "bbb": ["ccc"]})

  def test_double(self):
    testdata = """
    # test
    ## testb
    ### testc
    :: aaa
    bbb
    :: bbb
    ccc
    ## testd
    ### teste
    :: aaa
    ddd
    :: bbb
    eee"""
    data = main.generate_testlist(testdata)
    self.assertEqual(len(data), 2)
    self.assertEqual(data[0]["items"], ["test", "testb", "testc"])
    self.assertEqual(data[0]["exams"], {"aaa": ["bbb"], "bbb": ["ccc"]})
    self.assertEqual(data[1]["items"], ["test", "testd", "teste"])
    self.assertEqual(data[1]["exams"], {"aaa": ["ddd"], "bbb": ["eee"]})

  def test_triple(self):
    testdata = """
    # test
    ## testb
    ### testc
    :: aaa
    bbb
    :: bbb
    ccc
    ## testd
    ### teste
    :: aaa
    ddd
    :: bbb
    eee
    # testf
    ## testg
    ### testh
    :: aaa
    fff
    :: bbb
    ggg"""
    data = main.generate_testlist(testdata)
    self.assertEqual(len(data), 3)
    self.assertEqual(data[0]["items"], ["test", "testb", "testc"])
    self.assertEqual(data[0]["exams"], {"aaa": ["bbb"], "bbb": ["ccc"]})
    self.assertEqual(data[1]["items"], ["test", "testd", "teste"])
    self.assertEqual(data[1]["exams"], {"aaa": ["ddd"], "bbb": ["eee"]})
    self.assertEqual(data[2]["items"], ["testf", "testg", "testh"])
    self.assertEqual(data[2]["exams"], {"aaa": ["fff"], "bbb": ["ggg"]})

  def test_in_middle(self):
    testdata = """
    # test
    ## testb
    ### testc
    ### testd
    ## teste
    ### testf
    :: aaa
    bbb
    :: bbb
    ccc"""
    data = main.generate_testlist(testdata)
    self.assertEqual(len(data), 1)
    self.assertEqual(data[0]["items"], ["test", "teste", "testf"])
    self.assertEqual(data[0]["exams"], {"aaa": ["bbb"], "bbb": ["ccc"]})

  def test_ampersand(self):
    testdata = """
    # test
    ## testb
    ### testc
    :: aaa
    bbb
    :: bbb
    ccc
    ## teste
    ### testf
    :: aaa &&
    :: bbb
    ddd"""
    data = main.generate_testlist(testdata)
    self.assertEqual(len(data), 2)
    self.assertEqual(data[0]["items"], ["test", "testb", "testc"])
    self.assertEqual(data[0]["exams"], {"aaa": ["bbb"], "bbb": ["ccc"]})
    self.assertEqual(data[1]["items"], ["test", "teste", "testf"])
    self.assertEqual(data[1]["exams"], {"aaa": ["bbb"], "bbb": ["ddd"]})

  def test_in_middle_with_not_implemented(self):
    testdata = """
    # test
    ## testb
    ### testc
    ### testd
    :: aaa
    not implemented
    ## teste
    ### testf
    :: aaa
    bbb
    :: bbb
    ccc"""
    data = main.generate_testlist(testdata)
    self.assertEqual(len(data), 2)
    self.assertEqual(data[0]["items"], ["test", "testb", "testd"])
    self.assertEqual(data[0]["exams"], {"aaa": ["not implemented"]})
    self.assertEqual(data[1]["items"], ["test", "teste", "testf"])
    self.assertEqual(data[1]["exams"], {"aaa": ["bbb"], "bbb": ["ccc"]})

  def test_error_headers(self):
    testdata = """
    # test
    ## testb
    ### testc
    ##### testd
    :: aaa
    bbb
    :: bbb
    ccc"""
    with self.assertRaises(Exception):
      main.generate_testlist(testdata)

  def test_include(self):
    testdata = """
    # test
    ## testb
    ### testc
    :: aaa
    bbb
    :: bbb
    ccc
    &include({"name":"tests/test_include.md","test":"bbb"})"""
    data = main.generate_testlist(testdata)
    self.assertEqual(len(data), 2)
    self.assertEqual(data[0]["items"], ["test", "testb", "testc"])
    self.assertEqual(data[0]["exams"], {"aaa": ["bbb"], "bbb": ["ccc"]})
    self.assertEqual(data[1]["items"], ["test", "teste", "testf"])
    self.assertEqual(data[1]["exams"], {"aaa": ["bbb"], "bbb": ["ccc"]})

  def test_include_basedir(self):
    testdata = """
    # test
    ## testb
    ### testc
    :: aaa
    bbb
    :: bbb
    ccc
    &include({"name":"test_include.md","test":"bbb"})"""
    data = main.generate_testlist(testdata, "tests")
    self.assertEqual(len(data), 2)
    self.assertEqual(data[0]["items"], ["test", "testb", "testc"])
    self.assertEqual(data[0]["exams"], {"aaa": ["bbb"], "bbb": ["ccc"]})
    self.assertEqual(data[1]["items"], ["test", "teste", "testf"])
    self.assertEqual(data[1]["exams"], {"aaa": ["bbb"], "bbb": ["ccc"]})
