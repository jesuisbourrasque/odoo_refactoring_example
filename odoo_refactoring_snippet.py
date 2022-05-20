from odoo import models
import locale
from datetime import datetime
from odoo.tools import float_round
import json


class PartnerXlsx(models.AbstractModel):
    _name = "report.partner_book_xslx_report.partner_book_xslx_report"
    _inherit = "report.report_xlsx.abstract"
    _description = "Report xlsx"

    def generate_xlsx_report(self, workbook, data, partners):

        ############## Get nessesary values
        partner = data["partner_id"]
        debit_start_balance = self.env["account.move.line"].search(
            [
                ("display_type", "not in", ("line_section", "line_note")),
                ("partner_id", "=", partner),
                ("move_id.state", "=", "posted"),
                ("account_id.internal_type", "=", "receivable"),
                ("full_reconcile_id", "=", False),
                ("balance", "!=", 0),
                ("account_id.reconcile", "=", True),
                ("date", "<=", data["start_date"]),
            ]
        )
        if debit_start_balance:
            debit_start_balance = debit_start_balance.with_context(
                order_cumulated_balance="date desc, move_name desc, id, id"
            )[0].cumulated_balance
        else:
            debit_start_balance = 0
        credit_start_balance = self.env["account.move.line"].search(
            [
                ("display_type", "not in", ("line_section", "line_note")),
                ("partner_id", "=", partner),
                ("move_id.state", "=", "posted"),
                ("account_id.internal_type", "=", "payable"),
                ("full_reconcile_id", "=", False),
                ("balance", "!=", 0),
                ("account_id.reconcile", "=", True),
                ("date", "<=", data["start_date"]),
            ]
        )
        if credit_start_balance:
            credit_start_balance = credit_start_balance.with_context(
                order_cumulated_balance="date desc, move_name desc, id, id"
            )[0].cumulated_balance
        else:
            credit_start_balance = 0
        ##############

        locale.setlocale(locale.LC_ALL, "ru_RU.UTF-8")

        ############## Styles defenition
        sheet = workbook.add_worksheet("Акт сверки")
        bold = workbook.add_format({"bold": True})
        end_date_cell = workbook.add_format(
            {
                "align": "center",
                "valign": "vcenter",
                "bottom": 2,
            }
        )
        bold_centre = workbook.add_format(
            {
                "bold": True,
                "align": "center",
                "valign": "vcenter",
            }
        )
        italic = workbook.add_format(
            {
                "italic": True,
                "align": "center",
                "valign": "vcenter",
                "font_name": "Arial",
                "font_size": 12,
            }
        )
        saldo_cell = workbook.add_format(
            {
                "italic": True,
                "align": "center",
                "valign": "vcenter",
                "bottom": 2,
            }
        )
        saldo_cell_right = workbook.add_format(
            {
                "italic": True,
                "align": "center",
                "valign": "vcenter",
                "bottom": 2,
                "right": 2,
            }
        )
        debit_merge = workbook.add_format(
            {
                "bold": True,
                "align": "center",
                "valign": "vcenter",
                "top": 2,
                "left": 2,
                "right": 2,
                "bottom": 1,
            }
        )
        blue_underlined = workbook.add_format(
            {
                "align": "center",
                "valign": "vcenter",
                "font_color": "blue",
                "underline": True,
            }
        )
        centre = workbook.add_format(
            {
                "align": "center",
                "valign": "vcenter",
            }
        )
        right_align = workbook.add_format(
            {
                "align": "right",
                "valign": "vright",
            }
        )
        cell_format = workbook.add_format(
            {
                "font_name": "Arial",
                "font_size": 12,
                "align": "center",
                "valign": "vcenter",
                "border": 1,
            }
        )
        byn_cell = workbook.add_format(
            {
                "font_name": "Arial",
                "font_size": 12,
                "align": "center",
                "valign": "vcenter",
                "right": 2,
                "top": 1,
                "left": 1,
                "bottom": 1,
            }
        )
        empty_bottom_cells = workbook.add_format({"bottom": 2})
        empty_bottom_right_cells = workbook.add_format({"bottom": 2, "right": 2})
        empty_right_cells = workbook.add_format({"right": 2})
        empty_left_cells = workbook.add_format({"left": 2})
        ##############

        right_cells = workbook.add_format(
            {
                "right": 2,
                "align": "right",
                "valign": "vright",
            }
        )
        sheet.merge_range(
            "A1:J1",
            f"АКТ сверки взаимных расчетов между {self.env.company.name}"
            f" и {data['parnter_name']} по состоянию на {data['end_date']}",
            bold,
        )
        start_date = datetime.strptime(data["start_date"], "%Y-%m-%d")
        sheet.merge_range("A2:E2", "ДЕБЕТ", debit_merge)
        sheet.merge_range("F2:J2", "КРЕДИТ", debit_merge)
        sheet.write(2, 0, "Дата", cell_format)
        sheet.write(
            3,
            0,
            f"{start_date.day:02d}.{start_date.month:02d}.{start_date.year}",
            centre,
        )

        usd = self.env["res.currency"].search([("name", "=", "USD")])
        byn_balance = usd._convert(
            debit_start_balance,
            self.env["res.currency"].browse(data["currency_id"]),
            self.env.user.company_id,
            start_date,
        )

        sheet.merge_range("B3:C3", "Операция", cell_format)
        sheet.write(3, 1, "Сальдо", italic)
        sheet.write(3, 3, f"{float(debit_start_balance)} $", right_align)
        sheet.write(3, 4, f"{byn_balance} Br", right_cells)
        sheet.write(2, 1, "", cell_format)
        sheet.write(2, 7, "", cell_format)
        # sheet.write(3, 4, "", empty_right_cells)
        sheet.write(4, 4, "", empty_right_cells)
        sheet.write(3, 9, "", empty_right_cells)
        sheet.write(4, 9, "", empty_right_cells)
        sheet.write(2, 3, "Сумма", cell_format)
        sheet.write(2, 4, "Сумма в BYN", byn_cell)
        sheet.write(2, 5, "Дата", cell_format)
        sheet.write(
            3,
            5,
            f"{start_date.day:02d}.{start_date.month:02d}.{start_date.year}",
            centre,
        )

        usd = self.env["res.currency"].search([("name", "=", "USD")])
        byn_balance = usd._convert(
            credit_start_balance,
            self.env["res.currency"].browse(data["currency_id"]),
            self.env.user.company_id,
            start_date,
        )

        sheet.merge_range("G3:H3", "Операция", cell_format)
        sheet.write(3, 6, "Сальдо", italic)
        sheet.write(3, 8, f"{float(credit_start_balance)} $", right_align)
        sheet.write(3, 9, f"{byn_balance} Br", right_cells)
        sheet.write(2, 8, "Сумма", cell_format)
        sheet.write(2, 9, "Сумма в BYN", byn_cell)
        base_url = (
            self.env["ir.config_parameter"].sudo().get_param("web.base.url").rstrip("/")
        )
        row = 5
        credit_rows = 5
        col = 0
        spare_between = 0
        pp_rows = 5
        inv_rows = 6
        pp_rows_increase = 0
        inv_rows_increase = 0
        g_cell = 7
        c_cell = 6
        sum_debit = 0
        sum_credit = 0
        byn = self.env["res.currency"].search(
            [("name", "=", "BYN")],
        )
        currency_wizard = self.env["res.currency"].search(
            [("name", "=", data["currency_name"])],
        )
        unique_invoices = []
        right_items = []
        for waybill in reversed(data["waybills"]):
            sale_order = self.env["sale.order"].search(
                [("id", "=", waybill["sale_order_id"][0])]
            )
            not_unique_invoices = sale_order.invoice_ids.filtered(
                lambda x: x.amount_total > 0
            )
            [
                right_items.append(invoice)
                for invoice in not_unique_invoices
                if invoice not in unique_invoices
            ]

        right_column_items = sum(
            list(
                map(
                    lambda x: len(json.loads(x.invoice_payments_widget)["content"]),
                    right_items,
                )
            )
        )

        l = max(len(data["waybills"]), right_column_items) * 2 + 1

        for i in range(5, 5 + l):
            sheet.write(
                i,
                col,
                "",
                empty_left_cells,
            )
            sheet.write(
                i,
                col + 4,
                "",
                empty_right_cells,
            )
            sheet.write(
                i,
                col + 9,
                "",
                empty_right_cells,
            )

        for waybill in reversed(data["waybills"]):
            sale_order = self.env["sale.order"].search(
                [("id", "=", waybill["sale_order_id"][0])]
            )
            not_unique_invoices = sale_order.invoice_ids.filtered(
                lambda x: x.amount_total > 0
            )
            [
                unique_invoices.append(invoice)
                for invoice in not_unique_invoices
                if invoice not in unique_invoices
            ]

            waybill_date = datetime.strptime(waybill["date_waybill"], "%Y-%m-%d")
            sheet.write(
                row + spare_between,
                col,
                f"{waybill_date.day:02d}.{waybill_date.month:02d}.{waybill_date.year}",
                centre,
            )
            if waybill["bill_type_id"] == "tn":
                sheet.write(row + spare_between, col + 1, "ТН", centre)
            else:
                sheet.write(row + spare_between, col + 1, "ТТН", centre)
            link_waybill = f'{base_url}/web#id={waybill["id"]}&action=425&model=stock.waybill&view_type=form&cids=&menu_id=220"'
            sheet.write_url(
                f"C{c_cell}",
                f"{link_waybill}",
                blue_underlined,
                string=f'{waybill["series"]} {waybill["number"]} (по счету {waybill["sale_order_id"][1]})',
            )
            sheet.write(
                row + spare_between,
                col + 3,
                f"{round(waybill['amount_total'],2)} $",
                right_align,
            )
            sheet.write(
                row + spare_between,
                col + 4,
                f"{round(waybill['amount_total_byn'], 2)} Br",
                right_cells,
            )
            sheet.write(
                row + spare_between + 1,
                col + 4,
                "",
                empty_right_cells,
            )
            spare_between += 1
            row += 1
            c_cell += 2
        for invoice in unique_invoices:
            payments_widget = invoice.invoice_payments_widget
            if payments_widget != "false":
                payments_details = json.loads(payments_widget)
                for i in range(len(payments_details["content"])):
                    payments = self.env["account.payment"].search(
                        [
                            (
                                "id",
                                "=",
                                payments_details["content"][i]["account_payment_id"],
                            ),
                        ]
                    )
                    rate_byn = self.env["res.currency.rate"].search(
                        [
                            ("currency_id", "=", byn.id),
                            ("name", "=", payments.date),
                        ]
                    )
                    total_byn = float_round(
                        float(payments_details["content"][i]["amount"]) * rate_byn.rate,
                        2,
                    )
                    date_payment = datetime.strptime(
                        payments_details["content"][i]["date"], "%Y-%m-%d"
                    )
                    sheet.write(
                        inv_rows + inv_rows_increase,
                        col + 5,
                        f"{date_payment.day:02d}.{date_payment.month:02d}.{date_payment.year}",
                        centre,
                    )
                    sheet.write(
                        pp_rows + pp_rows_increase,
                        col + 6,
                        f"ПП",
                        bold_centre,
                    )
                    link = f'{base_url}/web#id={invoice.id}&action=206&active_id=528&model=account.move&view_type=form&cids=&menu_id=174"'
                    sheet.write_url(
                        f"G{g_cell}",
                        f"{link}",
                        blue_underlined,
                        string=invoice.name,
                    )
                    sheet.write(
                        inv_rows + inv_rows_increase,
                        col + 7,
                        payments.name,
                        bold,
                    )
                    amount = float_round(
                        float(payments_details["content"][i]["amount"]), 2
                    )
                    sum_credit += amount
                    sheet.write(
                        inv_rows + inv_rows_increase,
                        col + 8,
                        f"{round(amount, 2)} {payments_details['content'][i]['currency']}",
                        right_align,
                    )
                    sheet.write(
                        inv_rows + inv_rows_increase,
                        col + 9,
                        f"{round(total_byn, 2)} Br",
                        right_cells,
                    )
                    sheet.write(
                        inv_rows + inv_rows_increase + 1,
                        col + 9,
                        "",
                        empty_right_cells,
                    )
                    credit_rows += 1
                    pp_rows_increase += 2
                    inv_rows_increase += 2
                    g_cell += 2

        last_row_index = max([row + spare_between, inv_rows + inv_rows_increase])

        end_date = datetime.strptime(data["end_date"], "%Y-%m-%d")
        sheet.write(
            last_row_index,
            col,
            f"{end_date.day:02d}.{end_date.month:02d}.{end_date.year}",
            end_date_cell,
        )
        sheet.write(
            last_row_index,
            col + 5,
            f"{end_date.day:02d}.{end_date.month:02d}.{end_date.year}",
            end_date_cell,
        )
        sheet.write(
            last_row_index,
            col + 2,
            "",
            empty_bottom_cells,
        )
        sheet.write(
            last_row_index,
            col + 3,
            "",
            empty_bottom_cells,
        )
        sheet.write(
            last_row_index,
            col + 4,
            "",
            empty_bottom_right_cells,
        )

        debit_end_balance = self.env["account.move.line"].search(
            [
                ("display_type", "not in", ("line_section", "line_note")),
                ("partner_id", "=", partner),
                ("move_id.state", "=", "posted"),
                ("account_id.internal_type", "=", "receivable"),
                ("full_reconcile_id", "=", False),
                ("balance", "!=", 0),
                ("account_id.reconcile", "=", True),
                ("date", "<=", data["end_date"]),
            ]
        )
        if debit_end_balance:
            debit_end_balance = debit_end_balance.with_context(
                order_cumulated_balance="date desc, move_name desc, id, id"
            )[0].cumulated_balance
        else:
            debit_end_balance = 0
        usd = self.env["res.currency"].search([("name", "=", "USD")])
        end_debit_byn_balance = usd._convert(
            debit_end_balance,
            self.env["res.currency"].browse(data["currency_id"]),
            self.env.user.company_id,
            end_date,
        )
        credit_end_balance = self.env["account.move.line"].search(
            [
                ("display_type", "not in", ("line_section", "line_note")),
                ("partner_id", "=", partner),
                ("move_id.state", "=", "posted"),
                ("account_id.internal_type", "=", "payable"),
                ("full_reconcile_id", "=", False),
                ("balance", "!=", 0),
                ("account_id.reconcile", "=", True),
                ("date", "<=", data["end_date"]),
            ]
        )
        if credit_end_balance:
            credit_end_balance = credit_end_balance.with_context(
                order_cumulated_balance="date desc, move_name desc, id, id"
            )[0].cumulated_balance
        else:
            credit_end_balance = 0
        end_credit_byn_balance = usd._convert(
            credit_end_balance,
            self.env["res.currency"].browse(data["currency_id"]),
            self.env.user.company_id,
            end_date,
        )

        sheet.write(last_row_index, col + 1, "Сальдо", saldo_cell)
        sheet.write(
            last_row_index, col + 3, f"{float(debit_end_balance)} $", saldo_cell
        )
        sheet.write(
            last_row_index, col + 4, f"{end_debit_byn_balance} Br", saldo_cell_right
        )

        sheet.write(last_row_index, col + 6, "Сальдо", saldo_cell)
        sheet.write(
            last_row_index, col + 8, f"{float(credit_end_balance)} $", saldo_cell
        )
        sheet.write(
            last_row_index,
            col + 9,
            f"{float(end_credit_byn_balance)} Br",
            saldo_cell_right,
        )
        sheet.write(
            last_row_index,
            col + 7,
            "",
            empty_bottom_cells,
        )
        sheet.set_column("A:A", 10)
        sheet.set_column("D:D", 10)
        sheet.set_column("E:E", 12.5)

        sheet.set_column("F:F", 10)
        sheet.set_column("I:I", 10)
        sheet.set_column("J:J", 12.5)
