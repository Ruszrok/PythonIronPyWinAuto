"Classes that wrap custom controls"
from __future__ import unicode_literals
from __future__ import absolute_import

__revision__ = "$Revision: 716 $"

import time
import ctypes
import re

from .. import win32defines, win32structures, win32functions
from .. import findbestmatch
from .. import clipboard
from ..RemoteMemoryBlock import RemoteMemoryBlock
from . import HwndWrapper, win32_controls
from .. import SendKeysCtypes as SendKeys
from PIL import ImageGrab

from ..timings import Timings

# special SheetsWnd stuff
SVM_FIRST = 0x1800
SVM_GETPAGEINFO = SVM_FIRST + 1

class SHEETS_PAGE_INFO(win32structures.Structure):
    _fields_ = [
        ('page', win32structures.HWND),
        ('visible', win32structures.BOOL),
        ('enable', win32structures.BOOL)
    ]

    def IsAccessible(self):
        return bool(self.visible and self.enable)
        
def GetSheetsWndPageInfo(sheetsWnd, page):
    "Returns sheetsWnd page's info"

    if isinstance(page, basestring):
        # find the string in the tab control
        texts = sheetsWnd.Texts()
        best_text = findbestmatch.find_best_match(page, texts, texts)
        page = texts.index(best_text) - 1
        
    if page >= sheetsWnd.TabCount() or page < 0:
            raise IndexError(
                "Only %d tabs available, you asked for tab %d (zero based)" % (
                sheetsWnd.TabCount(),
                page))

    remote_mem_info = RemoteMemoryBlock(sheetsWnd)
    info = SHEETS_PAGE_INFO()
    pInfo = remote_mem_info.Address()

    sheetsWnd.SendMessage(SVM_GETPAGEINFO, page, remote_mem_info)

    remote_mem_info.Read(info, pInfo)
    del remote_mem_info

    return info

def IsSheetsWndPageAccessible(sheetsWnd, page):
    return GetSheetsWndPageInfo(sheetsWnd, page).IsAccessible()

def GetSheetsWndPage(sheetsWnd, page):
    return HwndWrapper.HwndWrapper(GetSheetsWndPageInfo(sheetsWnd, page).page)
    
def HasSheetsWndPageVisible(sheetsWnd, pageName):
    # find the string in the tab control
    texts = sheetsWnd.Texts()
    if pageName in texts[1:len(texts)]:
        page = texts.index(pageName)
        return bool(GetSheetsWndPageInfo(sheetsWnd, page - 1).visible)
    else:
        return False
    
# GridWnd messages
GVM_FIRST = 0x1800
GVM_GETROWCOUNT = GVM_FIRST + 1
GVM_GETCOLUMNCOUNT = GVM_FIRST + 2
GVM_GETCELLRECT = GVM_FIRST + 3
GVM_GETCELLVIEWRECT = GVM_FIRST + 4
GVM_MAKECELLVISIBLE = GVM_FIRST + 5
GVM_ISCELLVISIBLE = GVM_FIRST + 6
GVM_GETCELLVIEWNAMES = GVM_FIRST + 7
GVM_GETVIEWNAMES = GVM_FIRST + 8
GVM_CELLDATA = GVM_FIRST + 9
GVM_GETCELLTEXT = GVM_FIRST + 10

# special message for testing Crash Handler
GVM_CRASHMESSAGE = GVM_FIRST + 99


class GridError(Exception):
    OK = 0
    GVE_INCORRECT_GRID_ID = -1
    GVE_INCORRECT_RANGE   = -2
    GVE_INVISIBLE_CELL    = -3
    GVE_SCROLLABLE_CELL   = -4
    GVE_INCORRECT_CELL_ID = -5
    GVE_INCORRECT_VIEW_ID = -6
    GVE_NO_DATA_SUPPLIER  = -7

    #----------------------------------------------------------------
    def __init__(self, description, return_code):
        Exception.__init__(self, description)
        self.return_code = return_code

class Area:
    FIXED = win32functions.MakeLong(0, 0)
    CLIENT = win32functions.MakeLong(1, 1)

    LEFT_TOP = win32functions.MakeLong(0, 0)
    TOP = win32functions.MakeLong(1, 0)
    RIGHT_TOP = win32functions.MakeLong(2, 0)

    LEFT = win32functions.MakeLong(0, 1)
    CENTER = win32functions.MakeLong(1, 1)
    RIGHT = win32functions.MakeLong(2, 1)

    LEFT_BOTTOM = win32functions.MakeLong(0, 2)
    BOTTOM = win32functions.MakeLong(1, 2)
    RIGHT_BOTTOM = win32functions.MakeLong(2, 2)

    __IDs__ = { 'fixed': FIXED,
                'client': CLIENT,

                'left_top': LEFT_TOP,
                'top': TOP,
                'right_top': RIGHT_TOP,

                'left': LEFT,
                'center': CENTER,
                'right': RIGHT,

                'left_bottom': LEFT_BOTTOM,
                'bottom': BOTTOM,
                'right_bottom': RIGHT_BOTTOM}
    
    @staticmethod
    def id(name):
        return Area.__IDs__[name.lower()]

class CELL(win32structures.Structure):
    _fields_ = [
        ('gridID', win32structures.DWORD),
        ('row',    win32structures.DWORD),
        ('column', win32structures.DWORD)
    ]

class CELL_VIEW(win32structures.Structure):
    _fields_ = [
        ('gridID', win32structures.DWORD),
        ('row',    win32structures.DWORD),
        ('column', win32structures.DWORD),
        ('view_id', win32structures.DWORD)
    ]

class CELL_DATA(win32structures.Structure):
    _fields_ = [
        ('name', win32structures.CHAR * 256),
        ('simple_value', win32structures.DWORD),
        ('value', win32structures.LPVOID)
    ]

#====================================================================
class GridWrapper(HwndWrapper.HwndWrapper):
    """Class that wraps Windows Grid common control

    This class derives from HwndWrapper - so has all the methods o
    that class also

    **see** HwndWrapper.HwndWrapper_

    .. _HwndWrapper.HwndWrapper: class-pywinauto.controls.HwndWrapper.HwndWrapper.html

    """

    friendlyclassname = "Grid"
    windowclasses = ["GridWnd"]

    #----------------------------------------------------------------
    def __init__(self, hwnd):
        "Initialise the instance"
        super(GridWrapper, self).__init__(hwnd)

        self.writable_props.extend([
            'RowCount',
            'ColumnCount'])
        self.views = {}

    #-----------------------------------------------------------
    def __checkReturnCode(self, ret):
        if (ret == GridError.GVE_INCORRECT_GRID_ID):
            raise GridError('Incorrect grid ID!', ret)
        elif (ret == GridError.GVE_INCORRECT_RANGE):
            raise GridError('Incorrect cells range!', ret)
        elif (ret == GridError.GVE_INVISIBLE_CELL):
            raise GridError('Cell is fully invisible!', ret)
        elif (ret == GridError.GVE_INCORRECT_CELL_ID):
            raise GridError('Incorrect cell ID!', ret)
        elif (ret == GridError.GVE_INCORRECT_VIEW_ID):
            raise GridError('Incorrect view id!', ret)
        elif (ret == GridError.GVE_NO_DATA_SUPPLIER):
            raise GridError('Incorrect data supplier!', ret)
        elif (ret < GridError.OK):
            raise ctypes.WinError()

    #-----------------------------------------------------------
    def __ViewDictFromText(self, text):
        views = {}
        view_infos = text.split("|")
        for view_info in view_infos:
            [view_name, view_id] = view_info.split("#")
            view_name = view_name.lstrip("class ")
            view_name = view_name.lstrip("Grid::View").lower()
            view_id = int(view_id)
            if view_name not in views:
                views[view_name] = [view_id]
            else:
                views[view_name].append(view_id)

        return views

    #-----------------------------------------------------------
    def ViewIDByName(self, view_dict, view_name):
        try:
            best_view = findbestmatch.find_best_match(view_name, view_dict.keys(), view_dict.keys())
            res = view_dict[best_view]
            if len(res) > 1:
                raise 0
            return view_dict[best_view][0]
        except:
            p = re.compile('(.*)#(.*)')
            r = p.match(view_name)
            if r:
                try:
                    view_name = r.groups()[0]
                    view_id = int(r.groups()[1])
                    best_view = findbestmatch.find_best_match(view_name, view_dict.keys(), view_dict.keys())
                    view_dict[best_view].index(view_id)
                    return view_id
                except:
                    return -1
            return -1;

    #-----------------------------------------------------------
    def ViewNames(self, area = 'client', refresh = False):
        """Return views"""

        if len(self.views) != 0 and refresh:
            self.views = {}

        if len(self.views) != 0:
            return self.views

        ret = self.SendMessage(GVM_GETVIEWNAMES, Area.id(area))
        self.__checkReturnCode(ret)
        if ret < 0:
            return self.views
            
        remote_mem_text = RemoteMemoryBlock(self, size=ret)
        pText = remote_mem_text.Address()
        ret = self.SendMessage(GVM_GETVIEWNAMES, Area.id(area), remote_mem_text)
        self.__checkReturnCode(ret)

        text = ctypes.create_string_buffer(ret)
        remote_mem_text.Read(text, pText)
        text = text.value
        if len(text) > ret:
            self.actions.log('=====================>   WARNING!!! len(text) > ret   <======================')

        del remote_mem_text

        self.views = self.__ViewDictFromText(text)
        return self.views

    #-----------------------------------------------------------
    def CellViewNames(self, row, column, area = 'client'):
        """ """
        remote_mem_cell = RemoteMemoryBlock(self)
        cell = CELL()
        cell.row = row
        cell.column = column
        cell.gridID = Area.id(area)
        remote_mem_cell.Write(cell)

        ret = self.SendMessage(GVM_GETCELLVIEWNAMES, remote_mem_cell)
        self.__checkReturnCode(ret)
        if ret < 0:
            return {}

        remote_mem_text = RemoteMemoryBlock(self, size=ret)
        pText = remote_mem_text.Address()
        ret = self.SendMessage(GVM_GETCELLVIEWNAMES, remote_mem_cell, remote_mem_text)
        self.__checkReturnCode(ret)

        text = ctypes.create_string_buffer(ret)
        remote_mem_text.Read(text, pText)
        text = text.value
        if len(text) > ret:
            self.actions.log('=====================>   WARNING!!! len(text) > ret   <======================')

        del remote_mem_cell
        del remote_mem_text

        return self.__ViewDictFromText(text)

    #-----------------------------------------------------------
    def CellData(self, data_name, row, column, area = 'client', simple_value = -1, data = None, view = None):
        if row < 0:
            row = self.RowCount() + row

        if column < 0:
            column = self.ColumnCount() + column

        if isinstance(view, basestring):
            view_dict = self.CellViewNames(row, column, area)
            view = self.ViewIDByName(view_dict, view)
            if view < 0:
                raise GridError('Cannot match view ' + str(view) + ' within ' + str(view_dict), GridError.GVE_INCORRECT_VIEW_ID)

        remote_mem_cell = RemoteMemoryBlock(self)
        cell = CELL_VIEW()
        cell.row = row
        cell.column = column
        cell.gridID = Area.id(area)
        if view:
            cell.view_id = view
        else:
            cell.view_id = -1
        
        remote_mem_cell.Write(cell)

        size = ctypes.sizeof(CELL_DATA)
        if data:
            size += ctypes.sizeof(data) + 5

        remote_mem_cell_data = RemoteMemoryBlock(self, size = size)
        cell_data = CELL_DATA()
        cell_data.name = data_name + '\0'
        cell_data.simple_value = simple_value
        if data:
            cell_data.value = remote_mem_cell_data.Address() + ctypes.sizeof(CELL_DATA) + 1
        else:
            cell_data.value = 0

        remote_mem_cell_data.Write(cell_data)
        
        if data:
            remote_mem_cell_data.Write(data, cell_data.value)
            
        
        ret = self.SendMessage(GVM_CELLDATA, remote_mem_cell, remote_mem_cell_data)
        self.__checkReturnCode(ret)

        remote_mem_cell_data.Read(cell_data)

        if data:
            remote_mem_cell_data.Read(data, cell_data.value)

        del remote_mem_cell
        del remote_mem_cell_data
            
        return win32structures.DWORD(cell_data.simple_value)

    #-----------------------------------------------------------
    def RowCount(self, area = 'client'):
        """Return the number of rows"""
        ret = self.SendMessage(GVM_GETROWCOUNT, Area.id(area))
        self.__checkReturnCode(ret)
        return ret

    #-----------------------------------------------------------
    def ColumnCount(self, area = 'client'):
        """Return the number of columns"""
        ret = self.SendMessage(GVM_GETCOLUMNCOUNT, Area.id(area))
        self.__checkReturnCode(ret)
        return ret
       
    #-----------------------------------------------------------
    def CellRect(self, row, column, area = 'client'):
        "Return cell's logic rectangle (scrolling may be needed)"
        remote_mem_cell = RemoteMemoryBlock(self)
        remote_mem_rect = RemoteMemoryBlock(self)
        cell = CELL()
        cell.row = row
        cell.column = column
        cell.gridID = Area.id(area)
        rect = win32structures.RECT()
        pRect = remote_mem_rect.Address()

        remote_mem_cell.Write(cell)

        ret = self.SendMessage(GVM_GETCELLRECT, remote_mem_cell, remote_mem_rect)
        self.__checkReturnCode(ret)

        remote_mem_cell.Read(cell)
        remote_mem_rect.Read(rect, pRect)
        del remote_mem_cell
        del remote_mem_rect

        return rect #(int(rect.left), int(rect.top), int(rect.right), int(rect.bottom))

    #-----------------------------------------------------------
    def CellViewRect(self, row, column, view, area = 'client'):
        "Returns view's logic rectangle"
        viewId = view
        if isinstance(viewId, basestring):
            view_dict = self.CellViewNames(row, column, area)
            viewId = self.ViewIDByName(view_dict, viewId)
            if viewId < 0:
                raise GridError('Cannot match view "' + str(view) + '" within ' + str(view_dict), GridError.GVE_INCORRECT_VIEW_ID)
        
        remote_mem_cell_view = RemoteMemoryBlock(self)
        remote_mem_rect = RemoteMemoryBlock(self)
        cell_view = CELL_VIEW()
        cell_view.row = row
        cell_view.column = column
        cell_view.gridID = Area.id(area)
        cell_view.view_id = viewId
        rect = win32structures.RECT()
        pRect = remote_mem_rect.Address()

        remote_mem_cell_view.Write(cell_view)

        ret = self.SendMessage(GVM_GETCELLVIEWRECT, remote_mem_cell_view, remote_mem_rect)
        self.__checkReturnCode(ret)

        remote_mem_cell_view.Read(cell_view)
        remote_mem_rect.Read(rect, pRect)

        del remote_mem_cell_view
        del remote_mem_rect

        return rect

    #-----------------------------------------------------------
    def ClickCell(self, row, column, area = 'client', double = False, use_log = True,
                      button='left', timeout=1.0, retry_step=0.1):
        "row and column are display coordinates"
        self.MakeCellVisible(row, column, area)
        begin_time = time.time()
        while not self.IsCellVisible(row, column, area) and time.time() - begin_time < timeout:
            time.sleep(retry_step)
        if not self.IsCellVisible(row, column, area):
            raise GridError('Cell ' + str((row, column)) + ' is invisible!', GridError.GVE_INVISIBLE_CELL)
        
        rect = self.CellRect(row, column, area)
        self.ClickInput(button, rect.mid_point(), double, use_log = use_log)
        self.actions.log('Clicked cell ' + str((row, column)) + ' by ' + button + ' mouse button')

    #-----------------------------------------------------------
    def ClickCellView(self, row, column, view, area = 'client', double = False, use_log = True,
                      button='left', timeout=1.0, retry_step=0.1):
        "row and column are display coordinates"
        self.MakeCellVisible(row, column, area)
        begin_time = time.time()
        while not self.IsCellVisible(row, column, area) and time.time() - begin_time < timeout:
            time.sleep(retry_step)
        if not self.IsCellVisible(row, column, area):
            raise GridError('Cell ' + str((row, column)) + ' is invisible!', GridError.GVE_INVISIBLE_CELL)
        
        rect = self.CellViewRect(row, column, view, area)
        self.ClickInput(button, rect.mid_point(), double, use_log = use_log)
        self.actions.log('Clicked cell ' + str((row, column)) + ' by ' + button + ' mouse button')

    #-----------------------------------------------------------
    def SaveAsImage(self, file_name):
        """Copies grid as image"""
        self.TypeKeys('^i')
        # wait until grid replies
        win32functions.SendMessage(self, win32defines.WM_GETTEXTLENGTH, 0, 0)
        im = ImageGrab.grabclipboard()
        if im:
            im.save(file_name)
        else:
            raise Exception('Cannot grab image from clipboard.')

    #-----------------------------------------------------------
    def CellsText(self, row_start, column_start, row_end, column_end,
                   area = 'client', max_cell_text_length=20):
        """Return cell text"""

        remote_mem_cells = RemoteMemoryBlock(self)
        whole_string_size = max((row_end - row_start + 1) * (column_end - column_start + 1) * max_cell_text_length + 1, 4096)
        if whole_string_size > 4096:
            print 'CellsText(): whole_buffer_size (max str length) = ', whole_string_size
        remote_mem_text = RemoteMemoryBlock(self, size=whole_string_size)

        class CELL_RANGE_RECT(win32structures.Structure):
            _fields_ = [
                ('gridID',       win32structures.DWORD),
                ('row_start',    win32structures.DWORD),
                ('column_start', win32structures.DWORD),
                ('row_end',      win32structures.DWORD),
                ('column_end',   win32structures.DWORD)
            ]

        cells = CELL_RANGE_RECT()
        cells.row_start    = row_start
        cells.column_start = column_start
        cells.row_end      = row_end
        cells.column_end   = column_end
        cells.gridID = Area.id(area)

        pText = remote_mem_text.Address()
        remote_mem_cells.Write(cells)

        ret = self.SendMessage(GVM_GETCELLTEXT, remote_mem_cells, remote_mem_text)
        self.__checkReturnCode(ret)

        text = ctypes.create_string_buffer(whole_string_size)
        remote_mem_text.Read(text, pText)
        text = text.value
        if len(text) > whole_string_size:
            self.actions.log('=====================>   WARNING!!! len(text) > whole_string_size   <======================')

        del remote_mem_cells
        del remote_mem_text

        return text

    #-----------------------------------------------------------
    def CellsTextAll(self):
        self.TypeKeys('^c')
        # wait until grid replies
        win32functions.SendMessage(self, win32defines.WM_GETTEXTLENGTH, 0, 0)
        # wait some time
        time.sleep(0.5)
        return clipboard.GetData()

    #-----------------------------------------------------------
    def Cells(self, area = 'client', max_cell_text_length=20):
        rows = self.CellsText(0, 0, self.RowCount(area)-1, self.ColumnCount(area)-1, area, max_cell_text_length=max_cell_text_length).split('\r\n')
        cells = []
        for row in rows:
            if row:
                cells.append(row.split('\t'))
        return cells

    #-----------------------------------------------------------
    def Row(self, row, area = 'client', max_cell_text_length=20):
        if isinstance(row, int):
            return self.CellsText(row, 0, row, self.ColumnCount(area)-1, area).split('\t')
        elif isinstance(row, basestring):
            if Area.id(area) == Area.TOP:
                captionArea = 'fixed'
            else:
                captionArea = 'left'
            caption = self.Column(0, captionArea, max_cell_text_length=max_cell_text_length)
            row = caption.index(row)
            return self.CellsText(row, 0, row, self.ColumnCount(area)-1, area, max_cell_text_length=max_cell_text_length).split('\t') #.rstrip('\t\r\n')
        else:
            raise TypeError('Incorrect type of the parameter row: "' + str(type(row)) + '". Expected: string or integer.')

    #-----------------------------------------------------------
    def Column(self, column, area = 'client', max_cell_text_length=20):
        if self.RowCount() == 0:
            raise GridError('Column(): The grid is empty!', GridError.GVE_INCORRECT_RANGE)
        if isinstance(column, int):
            return self.CellsText(0, column, self.RowCount(area)-1, column, area).split('\r\n') #.rstrip('\t\r\n')
        elif isinstance(column, basestring):
            if Area.id(area) == Area.LEFT:
                captionArea = 'fixed'
            else:
                captionArea = 'top'
            caption = self.CellsText(0, 0, 0, self.ColumnCount(captionArea)-1, captionArea, max_cell_text_length=max_cell_text_length).split('\t')
            #print 'grid caption = ', caption
            if not (column in caption):
                raise GridError('Column "' + str(column) + '" not found in caption "' + str(caption) + '"', GridError.GVE_INCORRECT_RANGE)
            column = caption.index(column)
            return self.CellsText(0, column, self.RowCount(area)-1, column, area, max_cell_text_length=max_cell_text_length).split('\r\n') #.rstrip('\t\r\n')
        else:
            raise TypeError('Incorrect type of the parameter column: "' + str(type(column)) + '". Expected: string or integer.')

    #-----------------------------------------------------------
    def MakeCellVisible(self, row, column, area = 'client'):
        "Scroll until the cell is visible"
        remote_mem_cell = RemoteMemoryBlock(self)
        cell = CELL()
        cell.row = row
        cell.column = column
        cell.gridID = Area.id(area)

        remote_mem_cell.Write(cell)

        ret = self.SendMessage(GVM_MAKECELLVISIBLE, remote_mem_cell)
        self.__checkReturnCode(ret)

        del remote_mem_cell
        time.sleep(0.005)
        return ret

    #-----------------------------------------------------------
    def IsCellVisible(self, row, column, area = 'client'):
        "Tests if client cell is visible or not"
        "row and column are display coordinates"
        remote_mem_cell = RemoteMemoryBlock(self)
        cell = CELL()
        cell.row = row
        cell.column = column
        cell.gridID = Area.id(area)

        remote_mem_cell.Write(cell)

        ret = self.SendMessage(GVM_ISCELLVISIBLE, remote_mem_cell)

        del remote_mem_cell
        return ret == 0

    #-----------------------------------------------------------
    def GetRowsByPattern(self, text_pattern, column, area = 'client'):
        result = []
        pattern = re.compile(text_pattern)
        for row in range(0, self.RowCount(area)):
            cellText = self.CellText(row, column, area)
            if pattern.search(cellText):
                result.append(row)

        return result

    #-----------------------------------------------------------
    def GetRowsByText(self, text, column, area = 'client'):
        result = []
        for row in range(0, self.RowCount(area)):
            if self.CellText(row, column, area) == text:
                result.append(row)

        return result

    #-----------------------------------------------------------
    def ClickCellViewByPattern(self, text_pattern, column, view, area = 'client', double = False, use_log = True,
                      button='left', timeout=1.0, retry_step=0.1):

        rows = self.GetRowsByPattern(text_pattern, column, area)
        if len(rows) != 1:
            raise Exception('There is no exact match "' + text_pattern + '" in grid. Matched rows are: ' + str(rows) + '.')

        self.ClickCellView(rows[0], column, view, area, double, use_log, button, timeout, retry_step)

    #-----------------------------------------------------------
    def ClickCellViewByText(self, text, column, view, area = 'client', double = False, use_log = True,
                      button='left', timeout=1.0, retry_step=0.1):

        rows = self.GetRowsByText(text, column, area)
        if len(rows) != 1:
            raise Exception('There is no exact match "' + text + '" in grid. Matched rows are: ' + str(rows) + '.')

        self.ClickCellView(rows[0], column, view, area, double, use_log, button, timeout, retry_step)

    #-----------------------------------------------------------
    def SortBy(self,text_pattern):
        pattern = re.compile(text_pattern)
        for area in ['fixed','top', 'right_top']:
            for column in range(0, self.ColumnCount(area)):
                cellText = self.CellText(0, column, area)
                if pattern.search(cellText):
                    self.ClickCellView(0, column, 'sort', area)
                    return
       
    #-----------------------------------------------------------
    # cells data
    #-----------------------------------------------------------
    def IsCellSelected(self, row, column, view = None):
        "Tests if client cell is selected or not"
        "row and column are absolute coordinates"
        return bool(self.CellData('selection', row, column, view = view).value)

    #-----------------------------------------------------------
    def IsCellSelectedByPattern(self, text_pattern, column, view = None):

        rows = self.GetRowsByPattern(text_pattern, column, 'client')
        if len(rows) != 1:
            raise Exception('There is no exact match "' + text_pattern + '" in grid. Matched rows are: ' + str(rows) + '.')

        return self.IsCellSelected(rows[0], column, view)

    #-----------------------------------------------------------
    def IsCellSelectedByText(self, text, column, view = None):

        rows = self.GetRowsByText(text, column, 'client')
        if len(rows) != 1:
            raise Exception('There is no exact match "' + text + '" in grid. Matched rows are: ' + str(rows) + '.')

        return self.IsCellSelected(rows[0], column, view)

    #-----------------------------------------------------------
    def SetSelection(self, row_start, column_start, row_end, column_end, exclude=False, exact=True, view = None):
        """Sets cells selection"""

        class CELL_RANGE_RECT(win32structures.Structure):
            _fields_ = [
                ('gridID',       win32structures.DWORD),
                ('row_start',    win32structures.DWORD),
                ('column_start', win32structures.DWORD),
                ('row_end',      win32structures.DWORD),
                ('column_end',   win32structures.DWORD),
                ('exclude_exact_mask', win32structures.DWORD)
            ]
        SELECTION_EXCLUDE = 0x00000100
        SELECTION_EXACT = 0x00000001

        cells = CELL_RANGE_RECT()
        cells.row_start    = row_start
        cells.column_start = column_start
        cells.row_end      = row_end
        cells.column_end   = column_end
        cells.gridID = Area.CLIENT

        exclude_exact_mask_int = 0
        if exclude:
            exclude_exact_mask_int = exclude_exact_mask_int | SELECTION_EXCLUDE
        if exact:
            exclude_exact_mask_int = exclude_exact_mask_int | SELECTION_EXACT
        cells.exclude_exact_mask = win32structures.UINT(exclude_exact_mask_int)
        
        ret = self.CellData('set_selection_range', row_start, column_start, data = cells, view = view)

        if row_start == row_end and column_start == column_end:
            if exclude:
                self.actions.log('Deselected cell ' + str((row_start, column_start)))
            else:
                if exact:
                    self.actions.log('Selected cell ' + str((row_start, column_start)))
                else:
                    self.actions.log('Added cell ' + str((row_start, column_start)) + ' to selection')
        else:
            if exclude:
                self.actions.log('Deselected cells range from ' + str((row_start, column_start)) + ' to ' + str((row_end, column_end)))
            else:
                if exact:
                    self.actions.log('Selected cells range from ' + str((row_start, column_start)) + ' to ' + str((row_end, column_end)))
                else:
                    self.actions.log('Added cells range from ' + str((row_start, column_start)) + ' to ' + str((row_end, column_end)) + ' to selection')
        return ret

    #-----------------------------------------------------------
    def SelectAll(self):
        self.TypeKeys('^a')
        # wait until grid replies
        win32functions.SendMessage(self, win32defines.WM_GETTEXTLENGTH, 0, 0)

    #-----------------------------------------------------------
    def SetSelectionRowsByPattern(self, patterns, column = 0):
        if not isinstance(patterns, list):
            patterns = [patterns]

        exactly = True
        for i in range(len(patterns)): 
            rows = self.GetRowsByPattern(patterns[i],column)
            for row in rows:
                self.SetSelection (row,column,row,column,False,exactly)
                exactly = False

    #-----------------------------------------------------------
    def SetSelectionRowsByText(self, texts, column = 0):
        if not isinstance(texts, list):
            texts = [texts]

        exactly = True
        for i in range(len(texts)): 
            rows = self.GetRowsByText(texts[i],column)
            for row in rows:
                self.SetSelection (row,column,row,column,False,exactly)
                exactly = False

    #-----------------------------------------------------------
    def SetSelectionRows(self, rows, column = 0):
        for i in range(len(rows)):
            self.SetSelection (rows[i],column,rows[i],column,False,i==0)
            
    #-----------------------------------------------------------
    def GetSelectionRows(self, column = 0):
        selection = []
        for i in range(self.RowCount()):
            if self.IsCellSelected(i,column): selection.append(i)
        return selection
        
    #-----------------------------------------------------------
    def CellColor(self, row, column, area = 'client', view = 'color'):
        "Return cell's color"
        color = self.CellData('color', row, column, area, view = view).value
        return (color & 0xFF, color >> 8 & 0xFF, color >> 16 & 0xFF)

    #-----------------------------------------------------------
    def IsCellChecked(self, row, column, area = 'client', view = None):
        '''Get cell check-box/radiobutton state'''
        return bool(self.CellData('check', row, column, area, view = view).value)

    #-----------------------------------------------------------
    def IsCellCheckedByText(self, text, column, area = 'client', view = None):

        rows = self.GetRowsByText(text, column, area)
        if len(rows) != 1:
            raise Exception('There is no exact match "' + text + '" in grid. Matched rows are: ' + str(rows) + '.')

        return self.IsCellChecked(rows[0], column, area, view)

    #-----------------------------------------------------------
    def CellText(self, row, column, area = 'client', view = None):
        "Return cell's text"
        length = self.CellData('text', row, column, area, view = view)
        text = ctypes.create_string_buffer(length.value+1)
        self.CellData('text', row, column, area, data = text, view = view)
        return text.value

    #-----------------------------------------------------------
    def CellImageListIndex(self, row, column, area = 'client', view = None):
        '''Get cell check-box/radiobutton state'''
        return int(self.CellData('image_list_index', row, column, area, view = view).value)


# WVL messages
WVM_FIRST = 0x1800
WVM_GET_SHELLS_COUNT = WVM_FIRST + 1
WVM_GET_SELECTED_SHELLS_COUNT = WVM_FIRST + 2
WVM_GET_SELECTED_SHELLS = WVM_FIRST + 3
WVM_SET_SELECTED_SHELLS = WVM_FIRST + 4
WVM_MAKE_SHELL_VISIBLE = WVM_FIRST + 5
WVM_GET_SHELL_RECT = WVM_FIRST + 6
WVM_GET_SHELL_TEXT = WVM_FIRST + 7
WVM_IS_SHELL_IN_PROGRESS = WVM_FIRST + 8
WVM_CAPTURE_IMAGE = WVM_FIRST + 9
WVM_GET_BLOCK_SELECTION = WVM_FIRST + 10
WVM_SET_BLOCK_SELECTION = WVM_FIRST + 11
WVM_IS_WND_IN_PROGRESS = WVM_FIRST + 12

#====================================================================
class WVLWrapper(HwndWrapper.HwndWrapper):
    """Class that wraps Windows WVL common control

    This class derives from HwndWrapper - so has all the methods of
    that class also

    **see** HwndWrapper.HwndWrapper_

    .. _HwndWrapper.HwndWrapper: class-pywinauto.controls.HwndWrapper.HwndWrapper.html

    """

    friendlyclassname = "WVL viewer"
    windowclasses = ["WVLCtrl"]

    #----------------------------------------------------------------
    def __init__(self, hwnd):
        "Initialise the instance"
        super(WVLWrapper, self).__init__(hwnd)


    #-----------------------------------------------------------
    def GetShellsCount(self):
        """Return the number of selected shells"""
        ret = self.SendMessage(WVM_GET_SHELLS_COUNT)
        return ret

    #-----------------------------------------------------------
    def GetSelectedShellsCount(self):
        """Return the number of selected shells"""
        ret = self.SendMessage(WVM_GET_SELECTED_SHELLS_COUNT)
        return ret

    #-----------------------------------------------------------
    def GetSelectedShells(self):
        """Returns list of indexes of selected shells"""

        count = self.GetSelectedShellsCount()
        if count == 0:
            return []
            
        class DYN_ARRAY(ctypes.Structure):
            _fields_ = [("value", ctypes.c_int * count)]

        shells = DYN_ARRAY()
        remote_mem_shells = RemoteMemoryBlock(self, ctypes.sizeof(ctypes.c_int)*count)
        pData = remote_mem_shells.Address()

        ret = self.SendMessage(WVM_GET_SELECTED_SHELLS, remote_mem_shells, count)
        if ret != 0:
            raise Exception('Failed WVM_GET_SELECTED_SHELLS message')
        remote_mem_shells.Read(shells, pData)
        
        shells_list = []
        for value in shells.value:
            shells_list.append(int(value))

        del remote_mem_shells
        
        return shells_list

    #-----------------------------------------------------------
    def SetSelectedShells(self, shells_list):
        """Sets list of indexes of selected shells"""

        if isinstance(shells_list, basestring):
            shells_list = self.GetShellsByPattern(shells_list)

        count = len(shells_list)

        if count!=0:
            class DYN_ARRAY(ctypes.Structure):
                _fields_ = [("value", ctypes.c_int * count)]

            shells = DYN_ARRAY()
            for value in shells_list:
                shells.value[shells_list.index(value)] = value

            remote_mem_shells = RemoteMemoryBlock(self, ctypes.sizeof(ctypes.c_int)*count)
            remote_mem_shells.Write(shells.value)
            ret = self.SendMessage(WVM_SET_SELECTED_SHELLS, remote_mem_shells, count)
            del remote_mem_shells
        else:
            ret = self.SendMessage(WVM_SET_SELECTED_SHELLS, 0, 0)

        if ret != 0:
            raise Exception('Failed WVM_SET_SELECTED_SHELLS message')

    #-----------------------------------------------------------
    def SetSelectedShellsAll(self):
        self.SetSelectedShells(range(self.GetShellsCount()))

    #-----------------------------------------------------------
    def MakeShellVisible(self, shell_index):
        """Makes shell visible"""

        if shell_index >= self.GetShellsCount():
            raise Exception('Invalid shell_index')

        ret = self.SendMessage(WVM_MAKE_SHELL_VISIBLE, shell_index)
        if ret != 0:
            raise Exception('Failed WVM_MAKE_SHELL_VISIBLE message')

    #-----------------------------------------------------------
    def GetShellRect(self, shell_index):
        """Returns shell's rectangle"""

        if shell_index >= self.GetShellsCount():
            raise Exception('Invalid shell_index')

        rect = win32structures.RECT()
        remote_mem_rect = RemoteMemoryBlock(self)
        pData = remote_mem_rect.Address()

        ret = self.SendMessage(WVM_GET_SHELL_RECT, shell_index, remote_mem_rect)
        if ret != 0:
            raise Exception('Failed WVM_GET_SHELL_RECT message')
        remote_mem_rect.Read(rect, pData)

        del remote_mem_rect

        return rect

    #-----------------------------------------------------------
    def GetShellText(self, shell_index):
        """Returns shell's text"""

        if shell_index >= self.GetShellsCount():
            raise Exception('Invalid shell_index')

        remote_mem_text = RemoteMemoryBlock(self)
        pText = remote_mem_text.Address()

        ret = self.SendMessage(WVM_GET_SHELL_TEXT, shell_index, remote_mem_text)
        if ret != 0:
            raise Exception('Failed WVM_GET_SHELL_TEXT message')

        text = ctypes.create_string_buffer(2000)
        remote_mem_text.Read(text, pText)
        text = text.value

        del remote_mem_text
        
        return text

    #-----------------------------------------------------------
    def Texts(self):
        """Returns shell's texts"""
        return [self.GetShellText(i) for i in xrange(self.GetShellsCount())]

    #-----------------------------------------------------------
    def IsShellInProgress(self, shell_index):
        """Returns True if shell is in progress"""
    
        ret = self.SendMessage(WVM_IS_SHELL_IN_PROGRESS, shell_index)
        if ret < 0:
            raise Exception('Failed WVM_IS_SHELL_IN_PROGRESS message')

        return ret

   #-----------------------------------------------------------
    def WaitShellProgressCompletes(self, shell_index, time_out = 600):
        """"""

        while self.IsShellInProgress(shell_index):
            if time_out <= 0:
                raise Exception('Waiting WVL shell progress exceeded')
            time.sleep(1)
            time_out = time_out - 1

    #-----------------------------------------------------------
    def IsWndInProgress(self):
        """Returns True if wnd is in progress"""
    
        ret = self.SendMessage(WVM_IS_WND_IN_PROGRESS)
        if ret < 0:
            raise Exception('Failed WVM_IS_WND_IN_PROGRESS message')

        return ret

   #-----------------------------------------------------------
    def WaitWndProgressCompletes(self, time_out = 600):
        """"""

        while self.IsWndInProgress():
            if time_out <= 0:
                raise Exception('Waiting WVL wnd progress exceeded')
            time.sleep(1)
            time_out = time_out - 1

    #-----------------------------------------------------------
    def GetShellsByPattern(self, text_pattern):
        """Returns list of shells"""

        shells = []
        pattern = re.compile(text_pattern)
        #self.actions.log('text_pattern = ' + str(text_pattern))
        for i in xrange(0, self.GetShellsCount()):
            shell_text = self.GetShellText(i)
            shell_text = shell_text.replace('\n', '\t')
            #self.actions.log('shell_text = "' + str(shell_text[:200]) + '"')
            if pattern.search(shell_text):
                shells.append(i)

        return shells

    #-----------------------------------------------------------
    def CaptureImage(self, shell_index, resulution, file_path):
        """Saves image of the control to file"""
    
        class DYN_ARRAY_INT(ctypes.Structure):
            _fields_ = [("value", ctypes.c_int * 4)]

        resolution_array = DYN_ARRAY_INT()
        resolution_array.value[0] = shell_index
        resolution_array.value[1] = resulution[0]
        resolution_array.value[2] = resulution[1]
        # destination WVL_DS_TOOLTIP
        resolution_array.value[3] = 4

        remote_mem_resolution = RemoteMemoryBlock(self)
        remote_mem_resolution.Write(resolution_array.value)

        text = ctypes.create_string_buffer(2000)
        text.value = file_path
        remote_mem_file_path = RemoteMemoryBlock(self)
        remote_mem_file_path.Write(text)
        
        ret = self.SendMessage(WVM_CAPTURE_IMAGE, remote_mem_resolution, remote_mem_file_path)
        if ret != 0:
            raise Exception('Failed WVM_CAPTURE_IMAGE message')

        del remote_mem_resolution
        del remote_mem_file_path
     
    #-----------------------------------------------------------
    def GetBlockSelection(self, shell_index, reserved_blocks = 1000):
        """Gets list of selected blocks"""
    
        class DYN_ARRAY_INT(ctypes.Structure):
            _fields_ = [("value", ctypes.c_int * reserved_blocks)]

        int_array = DYN_ARRAY_INT()
        int_array.value[0] = reserved_blocks

        remote_mem_int = RemoteMemoryBlock(self)
        remote_mem_int.Write(int_array.value)
        pData = remote_mem_int.Address()

        ret = self.SendMessage(WVM_GET_BLOCK_SELECTION, shell_index, remote_mem_int)
        if ret != 0:
            raise Exception('Failed WVM_GET_BLOCK_SELECTION message')
        remote_mem_int.Read(int_array, pData)
        
        blocks_list = []
        blocks_size = int_array.value[0]
        for i in range(blocks_size):
            blocks_list.append((int(int_array.value[2 * i + 1]), int(int_array.value[2 * i + 2])))

        del remote_mem_int
        
        return blocks_list

    #-----------------------------------------------------------
    def SetBlockSelection(self, shell_index, blocks_to_select):
        """Sets list of selected blocks"""
    
        class DYN_ARRAY_INT(ctypes.Structure):
            _fields_ = [("value", ctypes.c_int * (len(2 * blocks_to_select) + 1))]

        int_array = DYN_ARRAY_INT()
        int_array.value[0] = len(blocks_to_select)
        for i in range(len(blocks_to_select)):
            int_array.value[2 * i + 1] = blocks_to_select[i][0]
            int_array.value[2 * i + 2] = blocks_to_select[i][1]

        remote_mem_int = RemoteMemoryBlock(self)
        remote_mem_int.Write(int_array.value)

        ret = self.SendMessage(WVM_SET_BLOCK_SELECTION, shell_index, remote_mem_int)
        if ret != 0:
            raise Exception('Failed WVM_SET_BLOCK_SELECTION message')

        del remote_mem_int

# IGL messages
IGL_TM_FIRST = 0x1800
IGL_TM_GET_PROGRESS = IGL_TM_FIRST + 1
IGL_TM_OBJ_TO_SCREEN = IGL_TM_FIRST + 2
IGL_TM_SCREEN_TO_OBJ = IGL_TM_FIRST + 3
IGL_TM_CAPTURE_IMAGE = IGL_TM_FIRST + 4
IGL_TM_RECT_ZOOM = IGL_TM_FIRST + 5
IGL_TM_RECT_SELECTION = IGL_TM_FIRST + 6
IGL_TM_SELECT_LEGEND_ITEM = IGL_TM_FIRST + 7
IGL_TM_GET_AXES_RANGES = IGL_TM_FIRST + 8

#====================================================================
class IGLWrapper(HwndWrapper.HwndWrapper):
    """Class that wraps Windows IGL control

    This class derives from HwndWrapper - so has all the methods of
    that class also

    **see** HwndWrapper.HwndWrapper_

    .. _HwndWrapper.HwndWrapper: class-pywinauto.controls.HwndWrapper.HwndWrapper.html

    """

    friendlyclassname = "IGL window"
    windowclasses = ["IGLWnd"]

    #----------------------------------------------------------------
    def __init__(self, hwnd):
        "Initialise the instance"
        super(IGLWrapper, self).__init__(hwnd)

    #-----------------------------------------------------------
    def GetProgress(self):
        """Returns progress"""
    
        remote_mem_float = RemoteMemoryBlock(self, 16)
        pText = remote_mem_float.Address()

        ret = self.SendMessage(IGL_TM_GET_PROGRESS, remote_mem_float)
        if ret != 0:
            raise Exception('Failed IGL_TM_GET_PROGRESS message')

        progress = ctypes.c_float(0)
        remote_mem_float.Read(progress, pText)

        del remote_mem_float
        
        return progress.value

   #-----------------------------------------------------------
    def WaitProgressCompletes(self, time_out = 6000, error_on_timeout = True, time_freeze = 60):
        """"""

        progress = self.GetProgress()
        progress_freeze = progress
        time_out_freeze = time_freeze

        print('Started waiting IGL progress... Initial timeout = ' + str(time_out))
        while progress != 1.0:
            if time_out <= 0 or time_out_freeze <= 0:
                if error_on_timeout:
                    raise Exception('Waiting IGL progress exceeded (time_out = ' + str(time_out) + '; time_out_freeze = ' + str(time_out_freeze) + ')')
                else:
                    return

            time.sleep(1)
            time_out -= 1

            progress = self.GetProgress()
            if progress != progress_freeze:
                progress_freeze = progress
                time_out_freeze = time_freeze
            else:
                time_out_freeze -= 1

    #-----------------------------------------------------------
    def ObjToClientCoords(self, obj_coords):
        """Returns client coordinates from object coordinates"""
    
        class DYN_ARRAY_DOUBLE(ctypes.Structure):
            _fields_ = [("value", ctypes.c_double * 2)]

        class DYN_ARRAY_INT(ctypes.Structure):
            _fields_ = [("value", ctypes.c_int * 2)]

        obj_coords_array = DYN_ARRAY_DOUBLE()
        obj_coords_array.value[0] = obj_coords[0]
        obj_coords_array.value[1] = obj_coords[1]

        remote_mem_obj = RemoteMemoryBlock(self)
        remote_mem_obj.Write(obj_coords_array.value)
        
        remote_mem_screen = RemoteMemoryBlock(self, 50)
        
        ret = self.SendMessage(IGL_TM_OBJ_TO_SCREEN, remote_mem_obj, remote_mem_screen)
        if ret != 0:
            raise Exception('Failed IGL_TM_OBJ_TO_SCREEN message')

        screen_coords_array = DYN_ARRAY_INT()
        remote_mem_screen.Read(screen_coords_array.value)
        
        result = (screen_coords_array.value[0], screen_coords_array.value[1])

        del remote_mem_obj
        del remote_mem_screen

        return result

    #-----------------------------------------------------------
    def ClientToObjCoords(self, client_coords):
        """Returns client coordinates from object coordinates"""
    
        class DYN_ARRAY_DOUBLE(ctypes.Structure):
            _fields_ = [("value", ctypes.c_double * 2)]

        class DYN_ARRAY_INT(ctypes.Structure):
            _fields_ = [("value", ctypes.c_int * 2)]

        screen_coords_array = DYN_ARRAY_INT()
        screen_coords_array.value[0] = client_coords[0]
        screen_coords_array.value[1] = client_coords[1]

        remote_mem_screen = RemoteMemoryBlock(self)
        remote_mem_screen.Write(screen_coords_array.value)
        
        remote_mem_obj = RemoteMemoryBlock(self, 50)
        
        ret = self.SendMessage(IGL_TM_SCREEN_TO_OBJ, remote_mem_screen, remote_mem_obj)
        if ret != 0:
            raise Exception('Failed IGL_TM_SCREEN_TO_OBJ message')

        obj_coords_array = DYN_ARRAY_DOUBLE()
        remote_mem_obj.Read(obj_coords_array.value)
        
        result = (obj_coords_array.value[0], obj_coords_array.value[1])

        del remote_mem_obj
        del remote_mem_screen

        return result

    #-----------------------------------------------------------
    def GetMouseObjCoords(self, wait_time = 0):
        """Returns obj coordinates of the mouse"""
        if wait_time > 0:
            time.sleep(wait_time)

        point = win32structures.POINT()
        win32functions.GetCursorPos(ctypes.byref(point))
        win32functions.ScreenToClient(self, ctypes.byref(point))
        return self.ClientToObjCoords((point.x, point.y))

    #-----------------------------------------------------------
    def CaptureImage(self, resulution, file_path):
        """Saves image of the control to file"""
    
        class DYN_ARRAY_INT(ctypes.Structure):
            _fields_ = [("value", ctypes.c_int * 2)]

        resolution_array = DYN_ARRAY_INT()
        resolution_array.value[0] = resulution[0]
        resolution_array.value[1] = resulution[1]

        remote_mem_resolution = RemoteMemoryBlock(self)
        remote_mem_resolution.Write(resolution_array.value)

        text = ctypes.create_string_buffer(file_path) #2000)
        #text.value = file_path
        remote_mem_file_path = RemoteMemoryBlock(self)
        remote_mem_file_path.Write(text)
        
        ret = self.SendMessage(IGL_TM_CAPTURE_IMAGE, remote_mem_resolution, remote_mem_file_path)
        
        del remote_mem_resolution
        del remote_mem_file_path
        if ret != 0:
            raise Exception('Failed IGL_TM_CAPTURE_IMAGE message')

    #-----------------------------------------------------------
    def ZoomRectangle(self, rectangle):
        """Zooms in object coordinates"""
        
        class DYN_ARRAY_DOUBLE(ctypes.Structure):
            _fields_ = [("value", ctypes.c_double * 4)]
        
        rectangle_array = DYN_ARRAY_DOUBLE()
        rectangle_array.value[0] = rectangle[0]
        rectangle_array.value[1] = rectangle[1]
        rectangle_array.value[2] = rectangle[2]
        rectangle_array.value[3] = rectangle[3]

        remote_mem_obj = RemoteMemoryBlock(self)
        remote_mem_obj.Write(rectangle_array.value)
       
        ret = self.SendMessage(IGL_TM_RECT_ZOOM, remote_mem_obj)
        
        del remote_mem_obj
        if ret != 0:
            raise Exception('Failed IGL_TM_RECT_ZOOM message (return code = ' + str(ret) + ')')

    #-----------------------------------------------------------
    def SelectRectangle(self, rectangle, is_ctrl_pressed = False):
        """Zooms in object coordinates"""
        
        class DYN_ARRAY_DOUBLE(ctypes.Structure):
            _fields_ = [("value", ctypes.c_double * 4)]
        
        rectangle_array = DYN_ARRAY_DOUBLE()
        rectangle_array.value[0] = rectangle[0]
        rectangle_array.value[1] = rectangle[1]
        rectangle_array.value[2] = rectangle[2]
        rectangle_array.value[3] = rectangle[3]

        remote_mem_obj = RemoteMemoryBlock(self)
        remote_mem_obj.Write(rectangle_array.value)
       
        ret = self.SendMessage(IGL_TM_RECT_SELECTION, remote_mem_obj, is_ctrl_pressed)
        
        del remote_mem_obj
        if ret != 0:
            raise Exception('Failed IGL_TM_RECT_SELECTION message')

    #-----------------------------------------------------------
    def SelectLegendItem(self, item, select=True):
        """Select specified legend item"""
        
        text = ctypes.create_string_buffer(len(item) + 1)
        text.value = item
        
        remote_mem_text = RemoteMemoryBlock(self, len(item) + 1)
        remote_mem_text.Write(text)
        
        remote_mem_flag = RemoteMemoryBlock(self)
        remote_mem_flag.Write(ctypes.c_bool(select))
       
        ret = self.SendMessage(IGL_TM_SELECT_LEGEND_ITEM, remote_mem_flag, remote_mem_text)
        
        del remote_mem_text
        del remote_mem_flag
        if ret != 0:
            raise Exception('Failed IGL_TM_SELECT_LEGEND_ITEM message')
        self.actions.log('Selected legend item "' + str(item) + '"')

   #-----------------------------------------------------------
    def GetAxesRanges(self):
        """"""
        max_len = 20000

        remote_mem_text = RemoteMemoryBlock(self, max_len + 1)
        pText = remote_mem_text.Address()

        ret = self.SendMessage(IGL_TM_GET_AXES_RANGES, remote_mem_text, max_len)
        if ret != 0:
            raise Exception('Failed IGL_TM_GET_AXES_RANGES message with return code ' + str(ret))

        text = ctypes.create_string_buffer(max_len)
        remote_mem_text.Read(text, pText)
        text = text.value

        del remote_mem_text
        
        return text

    #-----------------------------------------------------------
    def ApplyMousePath(self, mousePath, objectCoords = True):
        if len(mousePath) == 0:
            return

        if objectCoords:
            _mousePath = []
            for coords in mousePath:
                _mousePath.append(self.ObjToClientCoords(coords))
        else:
            _mousePath = mousePath

        self.PressMouseInput(coords=_mousePath[0])
        for i in range(1, len(_mousePath)-1):
            self.MoveMouse(coords=_mousePath[i])
        self.ReleaseMouseInput(coords=_mousePath[len(_mousePath)-1])

# Property Grid messages
PG_TM_FIRST = 0x1800
PG_TM_SET_PROPERTY_VALUES = PG_TM_FIRST + 1
PG_TM_GET_PROPERTY_VALUES = PG_TM_FIRST + 2
PG_TM_GET_PROPERTY_VALUE = PG_TM_FIRST + 3
PG_TM_GET_PROPERTY_ATTRIBUTES = PG_TM_FIRST + 4
PG_TM_GET_PROPERTY_RECTANGLE = PG_TM_FIRST + 5

#====================================================================
class PropertyGridWrapper(HwndWrapper.HwndWrapper):
    """Class that wraps Windows Property Grid control

    This class derives from HwndWrapper - so has all the methods of
    that class also

    **see** HwndWrapper.HwndWrapper_

    .. _HwndWrapper.HwndWrapper: class-pywinauto.controls.HwndWrapper.HwndWrapper.html

    """

    friendlyclassname = "Property Grid window"
    windowclasses = ["PropertyGridWnd"]

    __types__ = {'CPropertyUInt': int,
                 'CPropertyDataFilter': str,
                 'CPropertyBool': bool,
                 'CPropertyEnum': str,
                 'CPropertyFloat': float}

    #----------------------------------------------------------------
    def __init__(self, hwnd):
        "Initialize the instance"
        super(PropertyGridWrapper, self).__init__(hwnd)

    #-----------------------------------------------------------
    def SetPropertyValues(self, propertyValue):
        """Sets property value"""
        max_len = 20000
    
        text = ctypes.create_string_buffer(max_len + 1)
        text.value = propertyValue
        remote_mem_file_value = RemoteMemoryBlock(self, max_len + 1)
        remote_mem_file_value.Write(text)
        
        self.actions.log('Setting property    ' + str(propertyValue))
        ret = self.SendMessage(PG_TM_SET_PROPERTY_VALUES, remote_mem_file_value, max_len)
        if ret != 0:
            if ret == win32defines.S_FALSE:
                remote_mem_file_value.Read(text, remote_mem_file_value.Address())
                raise Exception('PropertyGrid.SetPropertyValues: ' + str(text.value))
            else:
                raise Exception('Failed PG_TM_SET_PROPERTY_VALUES message with return code ' + str(ret))

        del remote_mem_file_value

    #-----------------------------------------------------------
    def GetPropertyValues(self):
        """Gets property values"""
    
        max_len = 20000

        remote_mem_text = RemoteMemoryBlock(self, max_len + 1)
        pText = remote_mem_text.Address()

        ret = self.SendMessage(PG_TM_GET_PROPERTY_VALUES, remote_mem_text, max_len)
        if ret != 0:
            raise Exception('Failed PG_TM_GET_PROPERTY_VALUES message with return code ' + str(ret))

        text = ctypes.create_string_buffer(max_len)
        remote_mem_text.Read(text, pText)
        text = text.value

        del remote_mem_text
        
        return text

    #-----------------------------------------------------------
    def GetPropertyValue(self, propertyName):
        """Gets property value"""
    
        max_len = 2000

        text = ctypes.create_string_buffer(max_len + 1)
        text.value = propertyName

        remote_mem_text = RemoteMemoryBlock(self, max_len + 1)
        pText = remote_mem_text.Address()
        remote_mem_text.Write(text)

        ret = self.SendMessage(PG_TM_GET_PROPERTY_VALUE, remote_mem_text, max_len)
        if ret != 0:
            raise Exception('Failed PG_TM_GET_PROPERTY_VALUE message with return code ' + str(ret))

        text = ctypes.create_string_buffer(max_len + 1)
        remote_mem_text.Read(text, pText)
        text = text.value

        del remote_mem_text
        
        return text

    #-----------------------------------------------------------
    def GetPropertyAttributes(self, propertyName):
        """Gets property attributes
            the possible values
            "type" : property_type ["CPropertyBool", "CPropertyFloat" etc]
            "auto" : true/false
            "read_only" : true/false
            "values" : [possible enum/bool values]
            "min" : minimum numeric value
            "max" : maximum numeric value
            "step" : step for numeric value
        """

        max_len = 2000

        text = ctypes.create_string_buffer(max_len)
        text.value = propertyName

        remote_mem_text = RemoteMemoryBlock(self, max_len)
        pText = remote_mem_text.Address()
        remote_mem_text.Write(text)

        ret = self.SendMessage(PG_TM_GET_PROPERTY_ATTRIBUTES, remote_mem_text, max_len)
        if ret != 0:
            raise Exception('Failed PG_TM_GET_PROPERTY_ATTRIBUTES message with return code ' + str(ret) + ' for "' + str(propertyName) + '"')

        text = ctypes.create_string_buffer(max_len)
        remote_mem_text.Read(text, pText)
        text = text.value

        del remote_mem_text
        
        dict = {}
        for key_value in text.split(';'):
            [key, value] = key_value.split('=');
            if value.startswith('[') and value.endswith(']'):
                value_list = []
                for sub_value in value[1:-1].split(','):
                    value_list.append(sub_value)
                dict[key] = value_list
            else:
                dict[key] = value

        return dict

    #-----------------------------------------------------------
    def GetPropertyRectangle(self, propertyName):
        max_len = 2000

        text = ctypes.create_string_buffer(max_len)
        text.value = propertyName

        remote_mem_text = RemoteMemoryBlock(self, max_len)
        pText = remote_mem_text.Address()
        remote_mem_text.Write(text)

        rect = win32structures.RECT()
        remote_mem_rect = RemoteMemoryBlock(self)
        pData = remote_mem_rect.Address()

        ret = self.SendMessage(PG_TM_GET_PROPERTY_RECTANGLE, remote_mem_text, remote_mem_rect)
        if ret != 0:
            raise Exception('Failed PG_TM_GET_PROPERTY_RECTANGLE message with return code ' + str(ret))

        remote_mem_rect.Read(rect, pData)

        del remote_mem_rect

        return rect

    #-----------------------------------------------------------
    def IsPropertyAuto(self, propertyName):
        attr = self.GetPropertyAttributes(propertyName)
        if not attr.has_key('auto'):
            raise Exception('Property ' + str(propertyName) + ' is not auto-property')
        return attr.has_key('read_only')
        
    #-----------------------------------------------------------
    def SetPropertyAuto(self, propertyName, makeAuto = True):
        if self.IsPropertyAuto(propertyName) != makeAuto:
            rect = self.GetPropertyRectangle(propertyName)
            self.ClickInput(coords = (rect.right - 10, rect.top + 10))

    #-----------------------------------------------------------
    def StartInplaceEditProperty(self, propertyName, ctrlClass = None):
        self.ClickInput(coords = self.GetPropertyRectangle(propertyName).mid_point())
        time.sleep(0.5)
        controls = []
        for w in self.Children():
            props = w.GetProperties()
            if not ctrlClass or props['Class'] == ctrlClass:
                controls.append(w)
        if len(controls) == 0:
            raise Exception('No inplace controls to start property editing')
        elif len(controls) == 1:
            return controls[0]
        else:
            return controls
        
    #-----------------------------------------------------------
    def DebugInplaceEditProperty(self, propertyName):
        self.SetFocus()
        self.ClickInput(coords = self.GetPropertyRectangle(propertyName).mid_point())
        for w in self.Children():
            print w.Class()

    #-----------------------------------------------------------
    def InplaceEditPropertyByEdit(self, propertyName, keys = r'0'):
        edit = self.StartInplaceEditProperty(propertyName, 'Edit')
        edit.TypeKeys(keys = r'^a' + str(keys) + r'{ENTER}', with_spaces=True, set_foreground = False)
        
    #-----------------------------------------------------------
    def InplaceEditPropertyByUpDown(self, propertyName, delta = 1):
        upDown = self.StartInplaceEditProperty(propertyName, 'msctls_updown32')
        if delta > 0:
            for _ in range(delta):
                upDown.Increment()
        else:
            for _ in range(-delta):
                upDown.Decrement()
        upDown.TypeKeys(keys = r'{ENTER}', set_foreground = False)

    #-----------------------------------------------------------
    def InplaceEditPropertyByCombobox(self, propertyName, value = r'True'):
        combobox = self.StartInplaceEditProperty(propertyName, 'ComboBox')
        combobox.Select(value)
        if combobox.IsEnabled():
            combobox.TypeKeys(keys = r'{ENTER}', set_foreground = False)

    #-----------------------------------------------------------
    def InplaceEditPropertyByButton(self, propertyName, controlToReturn = 'Button'):
        bttn = self.StartInplaceEditProperty(propertyName, 'Button')
        bttn.ClickInput()
        if controlToReturn == 'Button':
            return bttn
        else:
            for w in self.Children():
                props = w.GetProperties()
                if props['Class'] == controlToReturn:
                    return w
            return None

    #-----------------------------------------------------------
    def IsPropertyAccessible(self, propertyName):

        dict = self.GetPropertyAttributes(propertyName)
        if not 'read_only' in dict or dict['read_only'] != 'true':
            return True
        else:
            return False

    #-----------------------------------------------------------
    def Dict(self, include_read_only=False):
        """Get Python dictionary representation"""
        
        def adjust_to_property_type(grid_type_name, value):
            if grid_type_name in self.__types__.keys():
                return self.__types__[grid_type_name](value)
            else:
                print 'Warning! Property type "', str(grid_type_name), '" cannot be recognized.'
                return str(value)
        
        res = {}
        print '--------------------\n\n', self.GetPropertyValues()
        for property_str in self.GetPropertyValues().split('\r\n'):
            #[key_value_str, type_str] = property_str.split(' // ')
            if not property_str.startswith('//'):
                [key, value] = property_str.split(' = ')
                attrs = self.GetPropertyAttributes(key)
                res[key] = adjust_to_property_type(attrs['type'], value)
            elif include_read_only:
                property_str = property_str.lstrip('//').rstrip('// property is not editable')
                [key, value] = property_str.split(' = ')
                #attrs = self.GetPropertyAttributes(key)
                res[key] = str(value) #adjust_to_property_type(attrs['type'], value)
        return res



# SplitWnd messages
SW_TM_FIRST = 0x1800
SW_TM_SET_PROPERTY_VALUES = SW_TM_FIRST + 1
SW_TM_GET_PROPERTY_VALUES = SW_TM_FIRST + 2
SW_TM_GET_PANELS = SW_TM_FIRST + 3

#====================================================================
class SplitWndWrapper(HwndWrapper.HwndWrapper):
    """Class that wraps Windows SplitWnd control

    This class derives from HwndWrapper - so has all the methods of
    that class also

    **see** HwndWrapper.HwndWrapper_

    .. _HwndWrapper.HwndWrapper: class-pywinauto.controls.HwndWrapper.HwndWrapper.html

    """

    friendlyclassname = "SplitWnd window"
    windowclasses = ["SplitWnd"]

    #----------------------------------------------------------------
    def __init__(self, hwnd):
        "Initialize the instance"
        super(SplitWndWrapper, self).__init__(hwnd)

    #-----------------------------------------------------------
    def SetPropertyValues(self, propetrtyValue):
        """Sets property value"""
    
        text = ctypes.create_string_buffer(2000)
        text.value = propetrtyValue
        remote_mem_file_value = RemoteMemoryBlock(self)
        remote_mem_file_value.Write(text)
        
        ret = self.SendMessage(SW_TM_SET_PROPERTY_VALUES, remote_mem_file_value)
        if ret != 0:
            raise Exception('Failed SW_TM_SET_PROPERTY_VALUES message with return code ' + str(ret))

        del remote_mem_file_value

    #-----------------------------------------------------------
    def GetPropertyValues(self):
        """Gets property values"""
        
        max_len = 2000
    
        remote_mem_text = RemoteMemoryBlock(self, max_len)
        pText = remote_mem_text.Address()

        ret = self.SendMessage(SW_TM_GET_PROPERTY_VALUES, remote_mem_text, max_len)
        if ret != 0:
            raise Exception('Failed SW_TM_GET_PROPERTY_VALUES message with return code ' + str(ret))

        text = ctypes.create_string_buffer(max_len)
        remote_mem_text.Read(text, pText)
        text = text.value

        del remote_mem_text
        
        return text

    #-----------------------------------------------------------
    def GetPanels(self):
        """Gets split sub-windows"""

        class DYN_ARRAY_HWND(ctypes.Structure):
            _fields_ = [("value", win32structures.HWND * 2)]

        remote_mem_values = RemoteMemoryBlock(self)
        pvalues = remote_mem_values.Address()

        ret = self.SendMessage(SW_TM_GET_PANELS, remote_mem_values)
        if ret != 0:
            raise Exception('Failed SW_TM_GET_PANELS message with return code ' + str(ret))

        panels = DYN_ARRAY_HWND()
        remote_mem_values.Read(panels.value)

        del remote_mem_values
        
        return (HwndWrapper.HwndWrapper(panels.value[0]), HwndWrapper.HwndWrapper(panels.value[1]))

# Ideal Split Tree messages
SPLITTREE_TM_FIRST = 0x1800
SPLITTREE_TM_CAPTURE_IMAGE = SPLITTREE_TM_FIRST + 1
SPLITTREE_TM_CONTEXT_MENU = SPLITTREE_TM_FIRST + 2

SPLITTREE_WRONG_PARAMETER = -1

#====================================================================
class SplitTreeWrapper(HwndWrapper.HwndWrapper):
    """Class that wraps Windows Ideal Split Tree control

    This class derives from HwndWrapper - so has all the methods of
    that class also

    **see** HwndWrapper.HwndWrapper_

    .. _HwndWrapper.HwndWrapper: class-pywinauto.controls.HwndWrapper.HwndWrapper.html

    """

    friendlyclassname = "SplitTree"
    windowclasses = ["IdealSplitTreeWnd"]

    #----------------------------------------------------------------
    def __init__(self, hwnd):
        "Initialize the instance"
        super(SplitTreeWrapper, self).__init__(hwnd)

    #-----------------------------------------------------------
    def CaptureImage(self, file_name):
        """Dump full tree image to the specified file. It doesn't depend on screen resolution."""
        
        self.MoveMouse() # to (0, 0) of the control
        
        remote_mem_text = RemoteMemoryBlock(self, len(file_name) + 1)
        text = ctypes.create_string_buffer(file_name)
        remote_mem_text.Write(text)

        ret = self.SendMessage(SPLITTREE_TM_CAPTURE_IMAGE, remote_mem_text)
        if ret != 0:
            raise Exception('Failed SPLITTREE_TM_CAPTURE_IMAGE message with return code ' + str(ret))

        del remote_mem_text
        return self

    #-----------------------------------------------------------
    def OpenContextMenu(self, node_string_id):
        """Open context menu for the specified node (use strings in the binary format: "01101" or empty string for root node)."""
        
        remote_mem_text = RemoteMemoryBlock(self, len(node_string_id) + 1)
        text = ctypes.create_string_buffer(node_string_id)
        remote_mem_text.Write(text)

        self.SendMessageTimeout(SPLITTREE_TM_CONTEXT_MENU, remote_mem_text.memAddress)
        #if ret != 0:
        #    raise Exception('Failed SPLITTREE_TM_CONTEXT_MENU message with return code ' + str(ret))

        del remote_mem_text
        return self


# IdealPRAVDA PatternWnd messages
PW_TM_FIRST = 0x1800
PW_TM_SET_PROPERTY_VALUES = PW_TM_FIRST + 1
PW_TM_GET_PROPERTY_VALUES = PW_TM_FIRST + 2
PW_TM_BIT_TO_CLIENT_RECT = PW_TM_FIRST + 3
PW_TM_ENSURE_VISIBLE_BITS = PW_TM_FIRST + 4

#====================================================================
class PatternWndWrapper(HwndWrapper.HwndWrapper):
    """Class that wraps Windows IdealPRAVDA PatternWnd control

    This class derives from HwndWrapper - so has all the methods of
    that class also

    **see** HwndWrapper.HwndWrapper_

    .. _HwndWrapper.HwndWrapper: class-pywinauto.controls.HwndWrapper.HwndWrapper.html

    """

    friendlyclassname = "PatternWnd"
    windowclasses = ["PatternWnd"]

    #----------------------------------------------------------------
    def __init__(self, hwnd):
        "Initialize the instance"
        super(PatternWndWrapper, self).__init__(hwnd)

    #-----------------------------------------------------------
    def CaptureImage(self, file_name):
        """Copies pattern control as image"""
        self.TypeKeys('^i')
        # wait until control replies
        win32functions.SendMessage(self, win32defines.WM_GETTEXTLENGTH, 0, 0)
        im = ImageGrab.grabclipboard()
        if im:
            im.save(file_name)
        else:
            raise Exception('Cannot grab image from clipboard.')

    #-----------------------------------------------------------
    def SetPropertyValues(self, propetrtyValue):
        """Sets property value"""
    
        text = ctypes.create_string_buffer(2000)
        text.value = propetrtyValue
        remote_mem_file_value = RemoteMemoryBlock(self)
        remote_mem_file_value.Write(text)
        
        ret = self.SendMessage(PW_TM_SET_PROPERTY_VALUES, remote_mem_file_value)
        if ret != 0:
            raise Exception('Failed PW_TM_SET_PROPERTY_VALUES message with return code ' + str(ret))

        del remote_mem_file_value

    #-----------------------------------------------------------
    def GetPropertyValues(self):
        """Gets property values"""
        
        max_len = 2000
    
        remote_mem_text = RemoteMemoryBlock(self, max_len)
        pText = remote_mem_text.Address()

        ret = self.SendMessage(PW_TM_GET_PROPERTY_VALUES, remote_mem_text, max_len)
        if ret != 0:
            raise Exception('Failed PW_TM_GET_PROPERTY_VALUES message with return code ' + str(ret))

        text = ctypes.create_string_buffer(max_len)
        remote_mem_text.Read(text, pText)
        text = text.value

        del remote_mem_text
        
        return text

    #-----------------------------------------------------------
    def BitToClientCoords(self, bit_coords):
        """Returns client coordinates from bit coordinates"""
    
        class DYN_ARRAY_INT(ctypes.Structure):
            _fields_ = [("value", ctypes.c_int * 2)]

        bit_coords_array = DYN_ARRAY_INT()
        bit_coords_array.value[0] = bit_coords[0]
        bit_coords_array.value[1] = bit_coords[1]

        remote_mem_bit = RemoteMemoryBlock(self)
        remote_mem_bit.Write(bit_coords_array.value)
        
        rect = win32structures.RECT()
        remote_mem_rect = RemoteMemoryBlock(self)
        pData = remote_mem_rect.Address()
        
        ret = self.SendMessage(PW_TM_BIT_TO_CLIENT_RECT, remote_mem_bit, remote_mem_rect)
        if ret != 0:
            raise Exception('Failed PW_TM_BIT_TO_CLIENT_RECT message with return code ' + str(ret))

        remote_mem_rect.Read(rect, pData)

        del remote_mem_bit
        del remote_mem_rect

        return rect

    #-----------------------------------------------------------
    def EnsureVisibleBits(self, bit1, bit2 = None):
        """Makes bits visible"""
    
        class DYN_ARRAY_INT(ctypes.Structure):
            _fields_ = [("value", ctypes.c_int * 2)]

        if not bit2:
            bit2 = bit1
            
        bit1_array = DYN_ARRAY_INT()
        bit1_array.value[0] = bit1[0]
        bit1_array.value[1] = bit1[1]

        bit2_array = DYN_ARRAY_INT()
        bit2_array.value[0] = bit2[0]
        bit2_array.value[1] = bit2[1]

        remote_mem_bit1 = RemoteMemoryBlock(self)
        remote_mem_bit1.Write(bit1_array.value)

        remote_mem_bit2 = RemoteMemoryBlock(self)
        remote_mem_bit2.Write(bit2_array.value)

        ret = self.SendMessage(PW_TM_ENSURE_VISIBLE_BITS, remote_mem_bit1, remote_mem_bit2)
        if ret != 0:
            raise Exception('Failed PW_TM_ENSURE_VISIBLE_BITS message with return code ' + str(ret))

        del remote_mem_bit1
        del remote_mem_bit2

    #-----------------------------------------------------------
    def ClickBit(self, bit, button = "left"):
        """Click to bit"""
        self.EnsureVisibleBits(bit)
        self.ClickInput(coords = self.BitToClientCoords(bit).mid_point(), button = button)

        
# IdealPRAVDA ComboBoxCategoryVariable messages
CBCV_TM_FIRST = 0x1800
CBCV_TM_GET_ITEM_CATEGORY = CBCV_TM_FIRST + 1

#====================================================================
class ComboBoxCategoryVariableWrapper(win32_controls.ComboBoxWrapper):
    """Class that wraps Windows IdealPRAVDA ComboBoxCategoryVariable control
    """

    friendlyclassname = "ComboBoxCategoryVariable"
    # no automatic window classe
    windowclasses = []

    #----------------------------------------------------------------
    def __init__(self, hwnd):
        "Initialize the instance"
        super(ComboBoxCategoryVariableWrapper, self).__init__(hwnd)

    #-----------------------------------------------------------
    def GetItemCategory(self, index):
        """Gets item category text value"""
        
        max_len = 2000
    
        remote_mem_text = RemoteMemoryBlock(self, max_len)
        pText = remote_mem_text.Address()

        ret = self.SendMessage(CBCV_TM_GET_ITEM_CATEGORY, index, remote_mem_text)
        if ret != 0:
            raise Exception('Failed CBCV_TM_GET_ITEM_CATEGORY message with return code ' + str(ret))

        text = ctypes.create_string_buffer(max_len)
        remote_mem_text.Read(text, pText)
        text = text.value

        del remote_mem_text
        
        return text

    #-----------------------------------------------------------
    def ItemCategories(self):
        """Returns list of categories"""

        categories = []
        for i in range(0, self.ItemCount()):
            categories.append(self.GetItemCategory(i))
        return categories

    #-----------------------------------------------------------
    def SelectCategory(self, category):
        """Selects combobox item by category"""

        categories = self.ItemCategories()
        best_category = findbestmatch.find_best_match(category, categories, categories)
        index = categories.index(best_category)
        self.Select(index)

    #-----------------------------------------------------------
    def SelectedCategory(self):
        """Gets selected category"""
        index = self.SelectedIndex()
        return self.GetItemCategory(index)
