import os
import zipfile
from typing import Never
from pathlib import Path

from .exceptions import ErrorPathToDocx, ErrorDocxFile


class DocxExtractor:
    def __init__(self, path: str | Path) -> None:
        self._path = path
        self._check_path()

    def _check_path(self) -> None | Never:  # type: ignore
        if not os.path.isdir(self._path):
            raise FileExistsError('Incorrect path to the file was passed!')


class DocxExtractorRead(DocxExtractor):
    def __init__(self, path: str | Path, name: str) -> None:
        """
        Args:
            path: path to the folder where
            the templates folder will be created.
            name: file name with docx extension.
        """

        super().__init__(path)
        self._name = name
        self._check_path_to_docx()
        self._check_zip()

    def read(self) -> None:
        path = os.path.join(self._path, 'templates')
        try:
            os.mkdir(path)
        except FileExistsError:
            # Skip the exception if the folder is created.
            pass
        with zipfile.ZipFile(os.path.join(self._path, self._name)) as zf:
            zf.extractall(path)

    def _check_path_to_docx(self) -> None | Never:  # type: ignore
        if not os.path.isfile(os.path.join(self._path, self._name)):
            raise ErrorPathToDocx(
                'The path to the file with docx extension is incorrect!'
            )

    def _check_zip(self) -> None | Never:  # type: ignore
        if not zipfile.is_zipfile(os.path.join(self._path, self._name)):
            raise ErrorDocxFile('Error! The file is corrupted!')


class DocxExtractorWrite(DocxExtractor):
    """
    Class responsible for adding template files
    to the archive with docx extension.
    """

    def __init__(
        self,
        path: str | Path,
        name_save: str = 'default.docx',
        path_save: str | Path | None = None,
    ) -> None:
        """
        Args:
            path: the name of the file to be saved with
            the docx extension.
            name: file name with docx extension.
            path_save: path to save a file with docx extension.
        """

        super().__init__(path)
        self._name = name_save
        self._path_save = path_save
        self._check_save_path()

    def write(self) -> None:
        """Packing templates into a file with docx extension."""

        if not self._chech_name_file():
            self._name = self._correct_name_file()
        with zipfile.ZipFile(
            self._get_path_docx(),
            mode='w',
            compression=zipfile.ZIP_DEFLATED,
        ) as zf:
            for root, _, files in os.walk(self._path):
                for file in files:
                    filepath = os.path.join(root, file)
                    arcname = os.path.relpath(filepath, self._path)
                    zf.write(filepath, arcname)

    def _correct_name_file(self) -> str:
        return f'{self._name}.docx'

    def _get_path_docx(self) -> str:
        if self._path_save and os.path.isdir(self._path_save):
            return os.path.join(self._path_save, self._name)
        return os.path.join(self._path, self._name)

    def _chech_name_file(self) -> bool:
        if 'docx' in self._name:
            return True
        return False

    def _check_save_path(self) -> None | Never:  # type: ignore
        if self._path_save and not os.path.isdir(self._path_save):
            raise FileExistsError('Incorrect file saving path passed!')
