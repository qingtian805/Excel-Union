"""Base Processor
处理 Excel 文件的基础类
提过各种与 Excel 文件交互的方法，包括创建工作表、读取和写入数据、保存工作簿等。
该基类设计为可由特定库（例如 openpyxl、xlrd、xlwt ）实现。

!: 该类是一个抽象基类，不能直接实例化，需要实现所有私有方法。
!: 实现类需要实现全部以 "_" 开头的方法

约束: 
1. 所有标记都是 int ，即序号。包括工作表、行、列等。
2. 不管底层实现如何，使用 0 作为第一个序号。
3. 开始与结束的实现应当与 Python 的切片一致。
4. 异常已经由基础类处理，无需在实现中处理。
"""

from typing import Union, Optional, Iterable

def _implement_(function_name: str):
    print("{}() is not implemented!".format(function_name))

class base_processor:
    def __init__(self,
                 book_path: str):
        self.book_path = book_path
        _implement_(self.__init__.__name__)

    def _save(self, book_path: str) -> bool:
        """保存工作簿到 book_path
        
        :return: 成功与否
        :rtype: bool
        """
        _implement_(self._save.__name__)
        
    def _get_sheets(self) -> Iterable[str]:
        """获取工作表的名称
        
        :return: 工作表的名称
        :rtype: Iterable[str]
        """
        _implement_(self._get_sheets.__name__)

    def _create_sheet(self,
                      sheet_name: str,
                      index: int) -> bool:
        """创建工作表
        
        :param sheet_name: 工作表的名称
        :type sheet_name: str
        :param index: 工作表的序号
        :type index: int
        :return: 成功与否
        :rtype: bool
        """
        _implement_(self._create_sheet.__name__)

    def _max_sheet(self) -> int:
        """获取工作簿的最大工作表数
        
        :return: 工作簿的最大工作表序号
        :rtype: int
        """
        _implement_(self._max_sheet.__name__)

    def _max_column(self, sheet_index) -> int:
        """获取工作表的最大列序号
        
        :param sheet_index: 工作表的序号
        :return: 工作表的最大列序号
        :rtype: int
        """
        _implement_(self._max_column.__name__)

    def _max_row(self, sheet_index) -> int:
        """获取工作表的最大行序号
        
        :param sheet_index: 工作表的序号
        :type sheet_index: int
        :return: 工作表的最大行序号
        :rtype: int
        """
        _implement_(self._max_row.__name__)

    def _get_row(self, sheet_index, row, start_col, end_col) -> Iterable:
        """获取指定工作表的某一行的数据
        
        :param sheet_index: 工作表的序号
        :type sheet_index: int
        :param row: 行数，使用数字作为行数
        :type row: int
        :param start_col: 开始获取的列数
        :type start_col: int
        :param end_col: 结束获取的列数
        :type end_col: int
        :return: 使用列表类型返回数据
        :rtype: list
        """
        _implement_(self._get_row.__name__)

    def _get_col(self, sheet_index, col, start_row, end_row) -> Iterable:
        """获取指定工作表的某一列的数据

        :param sheet_index: 工作表的序号
        :type sheet_index: int
        :param col: 列数，使用字母作为列数
        :type col: str
        :param start_row: 开始获取的行数
        :type start_row: int
        :param end_row: 结束获取的行数
        :type end_row: int
        :return: 使用列表类型返回数据
        :rtype: list
        """
        _implement_(self._get_col.__name__)

    def _write_row(self, sheet_index, row, data, start_col, end_col) -> bool:
        """将 data 写入指定工作表的某一行
        
        :param sheet_index: 工作表的序号
        :type sheet_index: int
        :param row: 行数，使用数字作为行数
        :type row: int
        :param data: 输入的数据
        :type data: Iterable
        :param start_col: 开始写入的列数
        :type start_col: int
        :param end_col: 结束写入的列数
        :type end_col: int
        :return: 说明
        :rtype: bool
        """
        _implement_(self._write_row.__name__)

    def _write_col(self, sheet_index, col, data, start_row, end_row) -> bool:
        """将 data 写入指定工作表的某一列
        :param sheet_index: 工作表的序号
        :type sheet_index: int
        :param col: 列数，使用字母作为列数
        :type col: str
        :param data: 输入的数据
        :type data: list
        :param start_row: 开始写入的行数，eg. 1, 2, 3
        :type start_row: int
        :param end_row: 结束写入的行数（包含），使用数字，eg. 1, 2, 3
        :type end_row: int
        :return: 说明
        :rtype: bool
        """
        _implement_(self._write_col.__name__)

    def get_sheets(self) -> Iterable[str]: 
        """获取工作簿中所有工作表的名称
        
        :return: 工作表的名称
        :rtype: typle[str]"""
        return self._get_sheets()

    def create_sheet(self,
                     sheet_name: str,
                     index: int) -> bool:
        """在工作簿中创建一个新的工作表
        
        :param sheet_name: 新工作表的名称
        :type sheet_name: str
        :param index: 新工作表的位置，支持逆向索引
        :type index: int
        :return: 成功与否
        :rtype: bool"""
        return self._create_sheet(sheet_name, index)

    def get_data(self,
                 sheet_index: int,
                 start_row: Optional[int] = None,
                 end_row: Optional[int] = None,
                 start_col: Optional[int] = None,
                 end_col: Optional[int] = None) -> list:
        """获取一张工作表的内容
        
        :param sheet_index: 工作表的序号
        :type sheet_index: int
        :param start_row: 开始获取的行数
        :type start_row: int
        :param end_row: 结束获取的行数
        :type end_row: int
        :param start_col: 开始获取的列数
        :type start_col: int
        :param end_col: 结束获取的列数
        :type end_col: int
        :return: 使用列表类型返回数据，如果出错，则返回 None
        :rtype: list | None
        """
        if start_col is None:
            start_col = 0
        if end_col is None:
            end_col = self._max_column(sheet_index)
        if start_row is None:
            start_row = 0
        if end_row is None:
            end_row = self._max_row(sheet_index)
        
        res = []
        for row in range(start_row, end_row):
            res.append(self._get_row(sheet_index, row, start_col, end_col))
        
        return res
        
    def get_row(self,
                sheet_index: int,
                row: int,
                start_col: Optional[int] = None,
                end_col: Optional[int] = None) -> list:
        """获取某一行的数据
        
        :param sheet_index: 工作表，可以使用名称或序号
        :type sheet_index: str | int
        :param row: 行数，使用数字作为行数
        :type row: int
        :param start: 开始获取的列数，使用字母作为列数
        :type start: str
        :param end: 结束的列数（包含），使用字母作为列数
        :type end: str
        :return: 使用列表类型返回数据
        :rtype: list
        """
        # 检查 sheet 和 row 的合法性
        if sheet_index < 0 or sheet_index > self._max_sheet():
            return []
        if row < 0 or row > self._max_row(sheet_index):
            return []

        if start_col is None:
            start_col = 0
        if end_col is None:
            end_col = self._max_column(sheet_index)

        return self._get_row(sheet_index, row, start_col, end_col)
    
    def get_col(self,
                sheet_index: int,
                col: str,
                start_row: Optional[int] = None,
                end_row: Optional[int] = None) -> Union[list, None]:
        """get_col 的 Docstring
        
        :param sheet_index: 工作表，可以使用名称或序号
        :type sheet_index: str | int
        :param col: 列数，使用字母作为列数
        :type col: str
        :param start: 开始获取的行数，eg. 1, 2, 3
        :type start: int
        :param end: 结束获取的行数（包含），使用数字，eg. 1, 2, 3
        :type end: int
        :return: 使用列表类型返回数据，如果出错，则返回 None
        :rtype: list | None
        """
        if col < 0 or col > self._max_column(sheet_index):
            return None
        if sheet_index < 0 or sheet_index > self._max_sheet():
            return None
        
        if start_row is None:
            start_row = 0
        elif start_row < 0:
            return None
        if end_row is None:
            end_row = self._max_row(sheet_index)
        elif end_row > self._max_row(sheet_index):
            return None
            
        return self._get_col(sheet_index, col, start_row, end_row)

    def write_data(self,
                   sheet_index: Union[int],
                   data: Iterable,
                   start_row: Optional[int] = None,
                   end_row: Optional[int] = None,
                   start_col: Optional[int] = None,
                   end_col: Optional[int] = None) -> bool:
        """写入数据到工作表
        
        :param sheet_index: 工作表，可以使用名称或序号指定
        :type sheet_index: 
        :param data: 要写入的数据
        :type data: Iterable
        :param start_row: 开始写入的行数，eg. 1, 2, 3
        :type start_row: int
        :param end_row: 结束写入的行数（包含），使用数字，eg. 1, 2, 3
        :type end_row: int
        :param start_col: 开始写入的列数，eg. "A", "B", "C"
        :type start_col: str
        :param end_col: 结束写入的列数（包含），eg. "A", "B", "C"
        :type end_col: str
        :return: 成功与否
        :rtype: bool"""
        if start_col is None:
            start_col = 0
        if end_col is None:
            end_col = max(*[len(i) for i in data])
        if start_row is None:
            start_row = 0
        if end_row is None:
            end_row = len(data)

        if start_col > end_col or start_row > end_row:
            return False
        
        if start_col < 0 or start_row < 0 or \
                end_col > self._max_column(sheet_index) or end_row > self._max_row(sheet_index):
            return False
        
        for i in range(len(data)):
            self._write_row(sheet_index, i + start_row, data[i], start_col, end_col)

    def write_row(self,
                  sheet_index: int,
                  row: int,
                  data: Iterable,
                  start_col: Optional[int] = None,
                  end_col: Optional[int] = None) -> bool:
        """将 data 写入指定工作表的某一行
        
        :param sheet_index: 工作表，可以使用名称或序号
        :type sheet_index: str | int
        :param row: 行数，使用数字作为行数
        :type row: int
        :param data: 输入的数据
        :type data: Iterable
        :param start: 开始写入的列数
        :type start: str
        :param end: 结束写入的列数（包含）
        :type end: str
        :return: 说明
        :rtype: bool
        """
        if start_col is None:
            start_col = 0
        if end_col is None:
            end_col = len(data)

        return self._write_row(sheet_index, row, data, start_col, end_col)  

    def write_col(self,
                  sheet_index: int,
                  col: str,
                  data: list,
                  start_row: Optional[int] = None,
                  end_row: Optional[int] = None) -> bool:
        """将 data 写入指定工作表的某一列

        :param sheet_index: 工作表，可以使用名称或序号
        :type sheet_index: str | int
        :param col: 列数，使用字母作为列数
        :type col: str
        :param data: 输入的数据
        :type data: list
        :param start: 开始写入的行数，eg. 1, 2, 3
        :type start: int
        :param end: 结束写入的行数（包含），使用数字，eg. 1, 2, 3
        :type end: int
        :return: 说明
        :rtype: bool
        """

        if start_row is None:
            start_row = 0
        if end_row is None:
            end_row = len(data)
        
        return self._write_col(sheet_index, col, data, start_row, end_row)
        
    def save(self,
             book_path: Optional[str] = None) -> bool:
        """保存工作簿，默认保存到原始文件，如果需要请指定 bookpath
        
        :param book_path: 工作簿的路径
        :type book_path: str | None
        :return: 说明
        :rtype: bool
        """
        if book_path is None:
            book_path = self.book_path
        return self._save(book_path)
       
if __name__ == "__main__":
    bp = base_processor("1")
    bp.get_sheets()
