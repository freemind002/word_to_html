import os
import shutil
from pathlib import Path, WindowsPath
from random import shuffle, uniform
from time import sleep
from typing import Any, Dict, List, Text

from bs4 import BeautifulSoup
from win32com import client


class WordToHtml(object):
    def __init__(self) -> None:
        self.word = client.Dispatch("Word.Application")
        self.old_doc = None

    def get_file_list(self, file_type: Text) -> List[WindowsPath]:
        """在指定的目錄下，取得所需類型的檔案資訊(檔名、路徑....)

        Args:
            file_type (Text): 使用此參數，決定路徑

        Returns:
            List[WindowsPath]: 返回包含所需類型檔案資訊的列表，每個檔案資訊包括檔名和路徑等相關資訊。
        """
        file_list = []
        if file_type == "doc":
            directory = Path(__file__).parent.joinpath("src", "word_data")
        elif file_type == "docx":
            directory = Path(__file__).parent.joinpath("src", "word_data")
        elif file_type == "docx_copy":
            directory = Path(__file__).parent.joinpath("src", "copy_file", "docx")
            file_type = "docx"
        elif file_type == "html_copy":
            directory = Path(__file__).parent.joinpath("src", "copy_file", "html")
            file_type = "html"

        for root, dirs, files in os.walk(directory):
            for file in files:
                if file.lower().endswith(file_type):
                    file_list.append(Path(root).joinpath(file).absolute())

        print(f"file_type為{file_type}，有 {len(file_list)} 筆")

        return file_list

    def doc_to_docx(self, doc_list: List[WindowsPath]):
        """將doc轉為docx，以利後面pandoc的使用

        Args:
            doc_list (List[WindowsPath]): 關於檔案格式為doc的資訊
        """
        for doc_index, doc in enumerate(doc_list):
            print(doc_index)
            print(doc)
            self.old_doc = self.word.Documents.Open(str(doc))
            self.old_doc.SaveAs(str(doc).replace(".doc", ".docx"), 16)
            self.old_doc.Close()
            self.old_doc = None

    def docx_to_html(self, docx_list: List[WindowsPath]):
        """將docx轉為html，以利後面beautifulsoup進行解析

        Args:
            docx_list (List[WindowsPath]): 關於檔案格式為docx的資訊
        """
        # 此處故意使用shuffle，因為要隨機抽取資料，看資料的格式
        shuffle(docx_list)
        for docx_index, docx in enumerate(docx_list):
            print(type(docx))
            print(docx_index)
            print(docx)
            path, filename = os.path.split(docx)
            print(filename)
            system_command = f"pandoc -o {str(docx).replace('docx', 'html').replace(' ', '')} {str(docx).replace(' ', '')}"
            print(system_command)
            os.system(system_command)
            sleep(uniform(1, 3))

    def html_to_sql(self, html_list: List[WindowsPath]):
        """將html檔案進行解析後，將資料放進sql中

        Args:
            html_list (List[WindowsPath]): 關於檔案格式為html的資訊
        """
        for html in html_list:
            path, filename = os.path.split(html)
            print(filename)
            fp = open(str(html), "r", encoding="utf-8")
            soup = BeautifulSoup(fp, "lxml")
            print(soup)
            sleep(uniform(2, 4))

    def copy_file(self, docx_list: List[WindowsPath], file_type: Text):
        """將docx的檔案進行複製

        Args:
            docx_list (List[WindowsPath]): 關於檔案格式為docx的資訊
            file_type (Text): 檔案的格式類型
        """
        for docx_index, docx in enumerate(docx_list):
            print(docx_index)
            print(docx)
            path, filename = os.path.split(docx)
            shutil.copyfile(
                str(docx),
                Path(__file__).parent.joinpath(
                    "src", "copy_file", file_type, str(filename)
                ),
            )

    def rename_docx(self, docx_list: List[WindowsPath]):
        """去除檔案名中的空格，避免後面使用pandoc轉換時出現錯誤

        Args:
            docx_list (List[WindowsPath]): 關於檔案格式為docx的資訊
        """
        for docx_index, docx in enumerate(docx_list):
            print(docx_index)
            path, filename = os.path.split(docx)
            os.rename(str(docx), str(docx).replace(" ", ""))

    def run_all(self):
        # 將doc轉成docx
        doc_list = self.get_file_list("doc")
        if doc_list:
            self.doc_to_docx(doc_list)
        # 複製docx
        docx_list = self.get_file_list("docx")
        if docx_list:
            self.copy_file(docx_list, "docx")
        # 去除docx的空格
        docx_list = self.get_file_list("docx_copy")
        if docx_list:
            self.rename_docx(docx_list)
        # 將docx轉成html
        docx_list = self.get_file_list("docx_copy")
        if docx_list:
            self.docx_to_html(docx_list)
        # 對html進行解析
        html_list = self.get_file_list("html_copy")
        if html_list:
            self.html_to_sql(html_list)

    def main(self):
        try:
            self.run_all()
        except Exception as e:
            print("發生錯誤！")
            print(e)
        else:
            print("程序順利執行完成！")
        finally:
            print("無論成功或失敗，我一定會執行！")
            if self.old_doc:
                self.old_doc.Close()
            if self.word:
                self.word.Quit()


if __name__ == "__main__":
    WordToHtml().main()
