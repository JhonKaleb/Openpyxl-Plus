from openpyxl.worksheet._write_only import WriteOnlyWorksheet
from openpyxl.utils.exceptions import ReadOnlyWorkbookException
from openpyxl.workbook import Workbook
from worksheet_plus import WorksheetPlus

class WorkbookPlus(Workbook):
    def __init__(self):
        """
        Overriding __init__ to create the first standard sheet as
        one of the WorksheetPlus type
        """
        super().__init__()
        self._sheets = []
        if not self.write_only:
            self._sheets.append(WorksheetPlus(self))

    def create_sheet(self, title=None, index=None):
        """Create a worksheet (at an optional index).

        Args:
            title (str): Optional title of the sheet.
            index (int): Optional position at which the sheet will be inserted

        Raises:
            ReadOnlyWorkbookException: If try to create an sheet in an Read Only wb.

        Returns:
            obj, WorksheetPlus: Returns an WorksheetPlus obj.
        """
        if self.read_only:
            raise ReadOnlyWorkbookException('Cannot create new sheet in a read-only workbook')

        if self.write_only :
            new_ws = WriteOnlyWorksheet(parent=self, title=title)
        else:
            new_ws = WorksheetPlus(parent=self, title=title)

        self._add_sheet(sheet=new_ws, index=index)
        return new_ws

    def set_header_in_all_sheets(self, array, merge_header_entire_row=True):
        """Set the first lines of the sheet with an array of rows, e.g:
        [["Sales in october 2019"]] or [["sales"]["jan","feb","mar"]]
        See doc for more usage example.

        Args:
            array (list): [description]
            merge_header_entire_row (bool, optional): [description]. Defaults to True.
        """
        for sheet_name in self.sheetnames:
            self[sheet_name].set_header(array, merge_header_entire_row)
