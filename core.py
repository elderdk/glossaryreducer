# -*- coding: utf-8 -*-
from openpyxl import load_workbook
import xlrd
import random
from collections import namedtuple
import string
from pathlib import Path


Segment = namedtuple("Segment", "row_num words".split())


class ExcelWorker:
    def __init__(self):
        pass

    def _avg_words(self, col: int) -> int:
        word_cnts = [
            len(str(seg.words).split(" "))
            for seg in self._segments(col)
            ]

        return sum(word_cnts) / len(word_cnts)

    def _sample(self, col):
        return random.sample(self._segments(col), 3)

    def _triage(self, col: int, n: int, col2: int) -> list:
        def remove_duplicates(lst: list) -> dict:
            return dict(lst)

        def wc_equal_or_less_than(words: str, n: int) -> bool:
            words = str(words).split(" ")
            return len(words) <= n and len(words) != 0

        source_words_list = [
            (str(seg.words).split(" ")[0], seg.row_num)
            for seg in self._segments(col)
            if wc_equal_or_less_than(seg.words, n)
        ]

        source_words_list = remove_duplicates(source_words_list)

        target_words_list = [
            self.ws.cell_value(val, col2) for val in source_words_list.values()
        ]

        term_tuples = list(zip(source_words_list, target_words_list))
        term_tuples = [
            tt
            for tt in term_tuples
            # remove term tuples where source and target terms are identical
            if tt[0] != tt[1]
            # remove single character terms after removing punctuation
            and len(
                tt[0].translate(str.maketrans("", "", string.punctuation))
                ) != 1
        ]

        return term_tuples


class OldExcel(ExcelWorker):
    def __init__(self, path: str) -> None:
        self.path = path
        self.wb = xlrd.open_workbook(path)
        self.ws = self.wb.sheet_by_index(0)
        self.filetype = "old"

    def _segments(self, col: int) -> list:
        return [
            Segment(row, self.ws.cell_value(row, col))
            for row in range(self.rows)
            ]

    @property
    def rows(self) -> int:
        return self.ws.nrows


class NewExcel(ExcelWorker):
    def __init__(self, path: str) -> None:
        self.wb = load_workbook(path)
        self.ws = self.wb.active
        self.filetype = "new"


class GlossaryReducer:
    """Excel glossary reducer based on wordcount

    This module takes an Excel file (both .xls and .xlsx) and reduces the
    number of glossary terms based on the source term's word count.

    Attributes:
        path (str): absolute path of the excel file

    """

    def __init__(self, path: str) -> None:

        self.path = path

        fi = Path(path)
        if fi.is_file() and fi.suffix == '.xlsx':
            self.wb = NewExcel(path)
        if fi.is_file() and fi.suffix == '.xls':
            self.wb = OldExcel(path)

        else:
            raise TypeError(
                "Could not recognize the file extension.\
                     Please provide an Excel file."
            )

    @property
    def rows(self) -> int:
        """int: number of rows that exist in the file"""
        return self.wb.rows

    @property
    def analyze(self) -> str:
        """Report generator"""
        pass

    def sample(self, col) -> str:
        """:obj:`list` of :obj:`Segment`: Shows samples of the specified column

        Args:
            col: index of column to get a sample for

        Returns:
            List of Segment namedtuples(row number, term)

        """
        return self.wb._sample(col)

    def avglen(self, col: int) -> float:
        """Returns avg word count on a given column

        Args:
            col: index of column to get a avg word count for.

        Return:
            Float of avg word count of the given column index

        """

        try:
            return self.wb._avg_words(col)
        except IndexError:
            print(f"Nothing in column {col}.")

    def triage(self, source_col: int, n: int, target_col: int) -> list:
        """Returns segments in col that match len(seg.split(' ')) =< n

        Also removes source term with character length of 1 after deleting
        any punctuation.

        Args:
            source_col: index of source column
            n: wordcount to triage the terms by
                e.g. n = 1 will triage all terms 2 words or longer
            target_col: index of target column

        Returns:
            List of tuples (source col term, target col term)

        """
        return self.wb._triage(source_col, n, target_col)

    def __str__(self):
        return "Reducer for " + self.path.split("/")[-1]


if __name__ == "__main__":
    path = r"C:\Users\admin\Desktop\KRAFTON Glossary.xls"
    gr = GlossaryReducer(path)
    a = gr.triage(3, 1, 1)
