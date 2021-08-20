import xlwings as xw
import datetime
from dataclasses import dataclass


@dataclass
class MenuItem:
    name: str  # 菜品名
    count: int  # 数量
    price: float  # 单价


@dataclass
class TicketData:
    id: str  # 单号
    number: str  # 台牌
    order: str  # 开单着
    cashier: str  # 收银员
    menu_list: list[MenuItem]  # 菜单列表
    date: datetime.datetime  # 时间


@dataclass
class RenderRange:
    start: tuple[int, int]  # 区间起始点
    end: tuple[int, int] = None  # 区间结束点
    render_format: str = "%s"  # 渲染模板


# 给 sheet 的 cell 赋值
def cell(sheet: xw.Sheet, render_range: RenderRange, value, number_format=None):
    sheet.range(render_range.start, render_range.end).value = render_range.render_format % value
    if number_format:
        sheet.range(render_range.start, render_range.end).number_format = number_format


# 在第 row 行前插入一行
def insert_row(sheet: xw.Sheet, row, col_range: tuple[int, int]):
    sheet.range((row, col_range[0]), (row, col_range[1])).insert()


class Template:

    def __init__(
            self,
            template_file: str,  # 模板文件路径
            id_range: RenderRange,
            number_range: RenderRange,
            order_range: RenderRange,
            cashier_range: RenderRange,
            menu_insert_row: int,  # 菜单开始 insert 行号
            key_range: RenderRange = None,  # key 字样 range
            print_time_range: RenderRange = None,  # 打印时间
    ):
        self.template_file = template_file
        self.id_range = id_range
        self.number_range = number_range
        self.order_range = order_range
        self.cashier_range = cashier_range
        self.menu_insert_row = menu_insert_row
        self.key_range = key_range
        self.print_time_range = print_time_range
        self.menu_start_col = 1

    def render(self, ticket_data: TicketData) -> xw.Book:
        template_wb = xw.Book(self.template_file)
        sheet = template_wb.sheets.active
        # 单号
        cell(sheet, self.id_range, ticket_data.id)
        # 牌号
        cell(sheet, self.number_range, ticket_data.number)
        # 开单人
        cell(sheet, self.order_range, ticket_data.order)
        # 收银
        cell(sheet, self.cashier_range, ticket_data.cashier)
        # key
        if self.key_range:
            cell(sheet, self.key_range, ticket_data.date.timestamp())
        # print time
        if self.print_time_range:
            cell(sheet, self.print_time_range, ticket_data.date.strftime("%Y-%m-%d %H:%M:%S"))

        # 插入菜单
        cur_row = self.menu_insert_row
        total_price = 0
        for menu_item in ticket_data.menu_list:
            price = menu_item.price * menu_item.count
            total_price += price
            insert_row(sheet, cur_row, (self.menu_start_col, self.menu_start_col + 2))
            cell(sheet, RenderRange((cur_row, self.menu_start_col)), menu_item.name)
            cell(sheet, RenderRange((cur_row, self.menu_start_col + 1)), menu_item.count)
            cell(sheet, RenderRange((cur_row, self.menu_start_col + 2)), price, number_format="0.00")
            cur_row += 1
        # 合计
        cell(sheet, RenderRange((cur_row, self.menu_start_col + 2)), total_price, number_format="0.00")
        return template_wb
