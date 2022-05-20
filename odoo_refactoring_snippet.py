import json
import locale
from datetime import datetime

from odoo.exceptions import UserError

from odoo import models


class PartnerXlsx(models.AbstractModel):
    _name = "report.partner_book_xslx_report.partner_book_xslx_report"
    _inherit = "report.report_xlsx.abstract"
    _description = "Report xlsx"

    def generate_xlsx_report(self, workbook, data, partners):
        locale.setlocale(locale.LC_ALL, "ru_RU.UTF-8")

        # CHECK WAYBILLS
        _data = {
            "data": data,
            "workbook": workbook,
            "sheet": workbook.add_worksheet("Акт сверки"),
            "start_date": datetime.strptime(data["start_date"], "%Y-%m-%d"),
            "end_date": datetime.strptime(data["end_date"], "%Y-%m-%d"),
            "c_line": 6,
        }

        waybills = sorted(
            filter(lambda x: x["state"] != "cancel", self.check_waybills(_data)),
            key=lambda x: x["date_waybill"]
        )

        self.print_act_head(_data)

        self.print_balances(_data, _data["start_date"])  # start balance

        self.print_waybills_and_invoices(_data, waybills)

        self.print_balances(_data, _data["end_date"])  # end balance

    def check_waybills(self, data):
        domain = [
            ("partner_id", "=", data["data"]["partner_id"]),
            ("date_waybill", ">=", data["start_date"]),
            ("date_waybill", "<=", data["end_date"]),
        ]
        waybills = self.env["stock.waybill"].search_read(domain)
        if waybills:
            return waybills
        raise UserError("Накладные отсутсвуют для данного контрагента")

    def print_act_head(self, data):
        bold_white_borders = data["workbook"].add_format(
            {
                "bold": True,
                "align": "center",
                "top": 1,
                "right": 1,
                "left": 1,
                "top_color": "#FFFFFF",  # White
                "right_color": "#FFFFFF",  # White
                "left_color": "#FFFFFF",  # White

            }
        )
        regular_center = data["workbook"].add_format(
            {
                "align": "center",
                "valign": "vcenter",
            }
        )

        data["sheet"].merge_range(
            "A1:L1",
            f"АКТ",
            bold_white_borders,
        )
        data["sheet"].merge_range(
            "A2:L2",
            f"сверки взаимных расчетов между",
            bold_white_borders,
        )
        data["sheet"].merge_range(
            "A3:L3",
            f"{self.env.company.name} и {data['data']['parnter_name']}",
            bold_white_borders,
        )
        data["sheet"].merge_range(
            "A4:L4",
            f"по состоянию на {data['end_date'].strftime('%d.%m.%Y')}",
            bold_white_borders,
        )

        data["sheet"].merge_range(
            "A5:F5",
            f"{self.env.company.name} {data['data']['currency_name']}",
            regular_center,
        )
        data["sheet"].merge_range(
            "G5:L5",
            f"{data['data']['parnter_name']} {data['data']['currency_name']}",
            regular_center,
        )

        data["sheet"].write("A6", "Дата", regular_center)
        data["sheet"].merge_range("B6:D6", "Операция/Документ", regular_center)
        data["sheet"].write("E6", "Дебет", regular_center)
        data["sheet"].write("F6", "Кредит", regular_center)

        data["sheet"].write("G6", "Дата", regular_center)
        data["sheet"].merge_range("H6:J6", "Операция/Документ", regular_center)
        data["sheet"].write("K6", "Дебет", regular_center)
        data["sheet"].write("L6", "Кредит", regular_center)

    def get_balance(self, data, internal_type, date):
        balance = self.env["account.move.line"].search(
            [
                ("display_type", "not in", ("line_section", "line_note")),
                ("partner_id", "=", data["data"]["partner_id"]),
                ("move_id.state", "=", "posted"),
                ("account_id.internal_type", "=", internal_type),
                ("full_reconcile_id", "=", False),
                ("balance", "!=", 0),
                ("account_id.reconcile", "=", True),
                ("date", "<=", date),
            ]
        )
        if balance:
            balance = balance.with_context(
                order_cumulated_balance="date desc, move_name desc, id, id"
            )[0].cumulated_balance
        else:
            balance = 0
        return balance

    def print_balances(self, data, date, ):

        regular_left = data["workbook"].add_format(
            {
                "align": "left",
                "valign": "vcenter",
            }
        )
        regular_right = data["workbook"].add_format(
            {
                "align": "right",
                "valign": "vcenter",
            }
        )

        debit_start_balance = self.get_balance(data, "receivable", date)

        credit_start_balance = self.get_balance(data, "payable", date)

        data["sheet"].write(data["c_line"], 0, date.strftime('%d.%m.%Y'), regular_left)
        data["sheet"].merge_range(data["c_line"], 1, data["c_line"], 3, "Сальдо", regular_left)

        data["sheet"].write(data["c_line"], 6, date.strftime('%d.%m.%Y'), regular_left)
        data["sheet"].merge_range(data["c_line"], 7, data["c_line"], 9, "Сальдо", regular_left)

        data["sheet"].write(data["c_line"], 4, debit_start_balance, regular_right)
        data["sheet"].write(data["c_line"], 5, credit_start_balance, regular_right)
        data["c_line"] += 1

    def print_waybills_and_invoices(self, data, waybills):
        regular_right = data["workbook"].add_format(
            {
                "align": "right",
                "valign": "vcenter",
            }
        )
        regular_left = data["workbook"].add_format(
            {
                "align": "left",
                "valign": "vcenter",
            }
        )
        # c MEANS OUR COMPANY
        c_date_row = 0  # A row
        c_operation_row = 1  # B row
        c_debit_row = 4  # E row
        c_credit_row = 5  # F row
        filtered_invoices = []
        payment_ids = []
        for waybill in waybills:
            data["sheet"].write(data["c_line"], c_date_row, waybill["date_waybill"].strftime('%d.%m.%Y'), regular_left)
            if waybill['bill_type_id'] == 'ttn':
                data["sheet"].merge_range(data["c_line"], c_operation_row, data["c_line"], c_operation_row + 2,
                                          f"TTH {waybill['number']}", regular_left)
            else:
                data["sheet"].merge_range(data["c_line"], c_operation_row, data["c_line"], c_operation_row + 2,
                                          f"TH {waybill['number']}", regular_left)
            data["sheet"].write(data["c_line"], c_credit_row, waybill["amount_total"], regular_right)
            data["c_line"] += 1

            sale_order = self.env["sale.order"].search(
                [("id", "=", waybill["sale_order_id"][0])]
            )
            not_filtered_invoices = sale_order.invoice_ids.filtered(
                lambda x: x.amount_total > 0
            )
            for invoice in not_filtered_invoices:
                if invoice not in filtered_invoices:
                    filtered_invoices.append(invoice)
                    payments_widget = invoice.invoice_payments_widget
                    if payments_widget != "false":
                        for payment in json.loads(payments_widget)["content"]:
                            payment_info = {
                                "move_id": payment["move_id"],
                                "amount": payment["amount"],
                                "date": payment["date"]
                            }
                            if payment_info not in payment_ids:
                                payment_ids.append(payment_info)
                                payment = self.env["account.move"].search(
                                    [
                                        (
                                            "id",
                                            "=",
                                            payment_info["move_id"],
                                        ),
                                    ]
                                )
                                data["sheet"].write(data["c_line"], c_date_row,
                                                    datetime.
                                                    strptime(payment_info["date"], "%Y-%m-%d").
                                                    strftime("%d.%m.%Y"),
                                                    regular_left
                                                    )
                                data["sheet"].merge_range(data["c_line"], c_operation_row, data["c_line"],
                                                          c_operation_row + 2, payment.name, regular_left
                                                          )
                                data["sheet"].write(data["c_line"], c_debit_row, payment_info["amount"], regular_right)
                                data["c_line"] += 1
