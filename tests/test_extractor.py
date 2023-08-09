'''
test_read_unzip_docx >> True
test_invalid_docx >> True
test_write_default_docx >> True
test_write_name_docx >> True
test_write_save_file >> True
test_invalid_path >> True
'''


import unittest
import os
import shutil
from pathlib import Path

from docxpy.core.extractor import DocxExtractorWrite, DocxExtractorRead
from docxpy.core.exceptions import ErrorPathToDocx, ErrorDocxFile

# Removed sorting of methods in tests.
unittest.TestLoader.sortTestMethodsUsing = None  # type: ignore

# Constants
DIRECTORY = Path(os.path.split(__file__)[0]) / 'templates'


class DocxExtractorTest(unittest.TestCase):
    """
    Tests to check the correct operation of
    the DocxExtractor class functional.
    """

    @classmethod
    def tearDownClass(cls) -> None:
        super().tearDownClass()
        shutil.rmtree(DIRECTORY)

    def test_read_unzip_docx(self) -> None:
        """Check unzipping of a file with docx extension."""
        path = Path(os.getcwd()) / 'tests'
        file_name = 'file.docx'
        self.assertTrue(
            (path / file_name).is_file(),
            'Error! No file: file.docx!',
        )
        DocxExtractorRead(path, file_name).read()
        self.assertTrue(
            (path / 'templates').is_dir(),
            'Error! The file won\'t unpack! Check DocxExtractorRead',
        )

    def test_invalid_docx(self) -> None:
        """
        Check for invalid data in DocxExtractorRead.
        """
        path = Path(os.getcwd()) / 'tests'
        file_name = 'document.docx'
        self.assertRaises(
            (ErrorPathToDocx, ErrorDocxFile),
            DocxExtractorRead,
            path=path,
            name=file_name,
        )

    def test_write_default_docx(self) -> None:
        """
        Check the creation of a file with
        the name default and the extension docx.
        """

        DocxExtractorWrite(DIRECTORY).write()
        self._check_file((DIRECTORY / 'default.docx').is_file())
        (DIRECTORY / 'default.docx').unlink(missing_ok=True)

    def test_write_name_docx(self) -> None:
        """Check the creation of a file with a given name."""

        name_file_docx = 'test.docx'
        DocxExtractorWrite(DIRECTORY, name_file_docx).write()
        self._check_file((DIRECTORY / name_file_docx).is_file())
        (DIRECTORY / name_file_docx).unlink(missing_ok=True)
        name_file = 'test'
        DocxExtractorWrite(DIRECTORY, name_file).write()
        self._check_file((DIRECTORY / name_file_docx).is_file())
        (DIRECTORY / name_file_docx).unlink(missing_ok=True)

    def test_write_save_file(self) -> None:
        """
        Checking that the file is created
        in the required location when you save it.
        """

        directory = Path(os.path.split(__file__)[0]) / 'save_dir'
        directory.mkdir(parents=True, exist_ok=True)
        DocxExtractorWrite(DIRECTORY, path_save=directory).write()
        self._check_file((directory / 'default.docx').is_file())
        shutil.rmtree(directory)

    def test_invalid_path(self) -> None:
        """Check the handling of an invalid path."""

        invalid_path = './test_dir/test_dir/@$'
        self.assertRaises(
            FileExistsError,
            DocxExtractorWrite,
            path=invalid_path,
        )
        self.assertRaises(
            FileExistsError,
            DocxExtractorWrite,
            path=DIRECTORY,
            path_save=invalid_path,
        )

    def _check_file(self, path_bool: bool) -> None:
        self.assertTrue(
            path_bool,
            'Error! The file is not being created! Check DocxExtractorWrite!',
        )
