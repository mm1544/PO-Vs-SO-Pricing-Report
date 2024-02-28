from odoo import models
from datetime import date, datetime, timedelta
from odoo.exceptions import UserError
import logging
import xlsxwriter
import io
import base64
import re

_logger = logging.getLogger(__name__)


class PricingReport(models.Model):
    _inherit = 'sale.order'

    HEADER_TEXT = 'PO vs SO Pricing Report'
    HEADER_VALUES_LIST = ['Sale Order', 'Purchase Order', 'Cost on Sale Order', 'Unit Price on Purchase Order',
                          'Quantity', 'Price Difference', 'Product Code', 'Product Name', 'Customer']

    def prepare_email_content(self):
        """Prepares the content of the email."""
        date_str = self.get_first_day_of_previous_month().strftime("%B %Y")
        return {
            'text_line_1': 'Hi,',
            'text_line_2': f'Please find attached the {self.HEADER_TEXT} for {date_str}.',
            'text_line_3': 'Kind regards,',
            'text_line_4': self.get_config_param('po_vs_so_pricing_report.email_company_name'),
            'table_width': 600,
        }

    def get_config_param(self, key):
        """Retrieves configuration parameters."""
        try:
            return self.env['ir.config_parameter'].get_param(key) or ''
        except Exception as e:
            _logger.error(f"Error getting configuration parameter {key}: {e}")
            return ''

    def get_first_day_of_previous_month(self):
        """Returns the first day of the previous month."""
        today = date.today()
        first_day_of_current_month = today.replace(day=1)
        first_day_of_previous_month = (
            first_day_of_current_month - timedelta(days=1)).replace(day=1)
        first_day_of_previous_month_datetime = datetime.combine(
            first_day_of_previous_month, datetime.min.time())

        return first_day_of_previous_month_datetime

    def get_last_day_of_previous_month(self):
        """Returns the last day of the previous month."""
        first_day_of_current_month = date.today().replace(day=1)
        last_day_of_previous_month = first_day_of_current_month - \
            timedelta(days=1)
        last_day_of_previous_month_datetime = datetime.combine(
            last_day_of_previous_month, datetime.max.time())

        return last_day_of_previous_month_datetime

    def get_sale_orders(self):
        """Finds Sale Orders within the previous month."""
        first_day = self.get_first_day_of_previous_month()
        last_day = self.get_last_day_of_previous_month()

        return self.env['sale.order'].search([
            ('date_order', '>=', first_day),
            ('date_order', '<=', last_day),
            ('state', 'in', ['sale', 'done'])
        ])

    def prepare_report_data(self, sale_orders):
        """Prepares the report data from sale orders."""
        res_list = []

        for sale_order in sale_orders:
            if not sale_order.order_line:
                continue

            for sale_line in sale_order.order_line:
                if not sale_line.purchase_line_ids:
                    continue

                for purchase_line in sale_line.purchase_line_ids:
                    note = ''
                    if purchase_line.order_id.state not in ['purchase', 'done']:
                        continue
                    if purchase_line.product_id.x_include_in_apple_s2w_report:
                        continue
                    if purchase_line.product_id.x_licence_length_months > 0:
                        continue
                    if purchase_line.product_id.type not in ["product"]:
                        continue

                    purchase_line_unit_price = purchase_line.price_unit
                    purchase_currency = purchase_line.order_id.currency_id
                    sale_currency = sale_line.order_id.currency_id

                    if purchase_currency != sale_currency:
                        note = f'SO currency is {sale_currency.name} and PO currency is {purchase_currency.name}'
                        sale_date = sale_order.date_order

                        # Convert purchase unit price to sale order's currency
                        purchase_line_unit_price = purchase_currency._convert(
                            purchase_line_unit_price, sale_currency, sale_order.company_id, sale_date)

                    price_difference = (
                        sale_line.purchase_price - purchase_line_unit_price) * purchase_line.product_qty

                    if price_difference > 0:
                        res_list.append([
                            sale_order.name,
                            purchase_line.order_id.name,
                            sale_line.purchase_price,
                            purchase_line_unit_price,
                            purchase_line.product_qty,
                            price_difference,
                            sale_line.product_id.default_code,
                            sale_line.product_id.name,
                            sale_order.partner_id.display_name,
                            note
                        ])

        return res_list

    def generate_xlsx_file(self, data_matrix):
        """Generates an XLSX file from the provided data."""
        # Create a new workbook using XlsxWriter
        buffer = io.BytesIO()
        workbook = xlsxwriter.Workbook(buffer, {'in_memory': True})

        # Defining a bold format for the header
        bold_format = workbook.add_format({'bold': True})

        # Your data and formatting logic goes here
        worksheet = workbook.add_worksheet()

        # Seting the width of the columns
        # Headers are in the first row of data_matrix and their length determines the column width
        for col_num, header in enumerate(self.HEADER_VALUES_LIST):
            column_width = len(header)
            if col_num in [6]:
                column_width += 15
            if col_num in [7, 8]:
                column_width += 40

            # Set the column width
            worksheet.set_column(col_num, col_num, column_width)
            worksheet.write(0, col_num, header, bold_format)

        # Write data to worksheet
        for row_num, row_data in enumerate(data_matrix, start=1):
            for col_num, cell_value in enumerate(row_data):
                if col_num == 9 and cell_value:
                    format_to_use = workbook.add_format(
                        {'bg_color': '#FFBF00',
                         # 'font_color': '#c47772',
                         })
                    if cell_value:
                        worksheet.set_column(col_num, col_num, len(cell_value))
                    worksheet.write(row_num, col_num,
                                    cell_value, format_to_use)
                else:
                    worksheet.write(row_num, col_num, cell_value)

        # Close the workbook to save changes
        workbook.close()

        # Get the binary data from the BytesIO buffer
        binary_data = buffer.getvalue()
        return base64.b64encode(binary_data)

    def generate_email_html(self, email_content):
        """Generates the HTML content for the email."""
        return f"""
        <!--?xml version="1.0"?-->
        <div style="background:#F0F0F0;color:#515166;padding:10px 0px;font-family:Arial,Helvetica,sans-serif;font-size:12px;">
            <table style="background-color:transparent;width:{email_content['table_width']}px;margin:5px auto;">
                <tbody>
                    <tr>
                        <td style="padding:0px;">
                            <a href="/" style="text-decoration-skip:objects;color:rgb(33, 183, 153);">
                                <img src="/web/binary/company_logo" style="border:0px;vertical-align: baseline; max-width: 100px; width: auto; height: auto;" class="o_we_selected_image" data-original-title="" title="" aria-describedby="tooltip935335">
                            </a>
                        </td>
                        <td style="padding:0px;text-align:right;vertical-align:middle;">&nbsp;</td>
                    </tr>
                </tbody>
            </table>
            <table style="background-color:transparent;width:{email_content['table_width']}px;margin:0px auto;background:white;border:1px solid #e1e1e1;">
                <tbody>
                    <tr>
                        <td style="padding:15px 20px 10px 20px;">
                            <p>{email_content['text_line_1']}</p>
                            </br>
                            <p>{email_content['text_line_2']}</p>
                            </br>
                            <p style="padding-top:20px;">{email_content['text_line_3']}</p>
                            <p>{email_content['text_line_4']}</p>
                        </td>
                    </tr>
                    <tr>
                        <td style="padding:15px 20px 10px 20px;">
                        </td>
                    </tr>
                </tbody>
            </table>
            <table style="background-color:transparent;width:{email_content['table_width']}px;margin:auto;text-align:center;font-size:12px;">
                <tbody>
                    <tr>
                        <td style="padding-top:10px;color:#afafaf;">
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>
        """

    def generate_email_attachment(self, binary_data, subject):
        """Generates an email attachment."""
        attachment_name = re.sub(r'[() /]', '_', f"{subject}.xlsx")
        return (attachment_name, binary_data)

    def log_message(self, message, function_name):
        """Logs a message with the specified level."""
        self.env['ir.logging'].create({
            'name': 'PO vs SO Pricing Report',
            'type': 'server',
            'dbname': self.env.cr.dbname,
            'level': 'info',
            'message': message,
            'path': __name__,
            'line': f'PricingReport.{function_name}',
            'func': function_name,
        })

    def send_email_with_attachment(self, subject, body, attachment):
        """Sends an email with the specified attachment."""
        try:
            mail_mail = self.env['mail.mail'].create({
                'email_to': self.get_config_param('po_vs_so_pricing_report.recipient_email'),
                'email_from': self.get_config_param('po_vs_so_pricing_report.sender_email'),
                'email_cc': self.get_config_param('po_vs_so_pricing_report.cc_email'),
                'reply_to': self.get_config_param('po_vs_so_pricing_report.reply_to_email'),
                'subject': subject,
                'body_html': body,
                'attachment_ids': [(0, 0, {'name': attachment[0], 'datas': attachment[1]})],
            })
            mail_mail.send()
            self.log_message('Email sent', 'send_email_with_attachment')

        except Exception as e:
            _logger.error(f"Error in sending email: {e}")

    def send_pricing_report(self):
        """Generates and sends the pricing report."""
        sale_orders = self.get_sale_orders()
        data_list = self.prepare_report_data(sale_orders)

        if not data_list:
            _logger.warning('No data to report.')
            return

        binary_data = self.generate_xlsx_file(data_list)
        subject = f"{self.HEADER_TEXT} ({date.today().strftime('%d/%m/%y')})"
        email_body = self.generate_email_html(self.prepare_email_content())
        attachment = self.generate_email_attachment(binary_data, subject)

        self.send_email_with_attachment(subject, email_body, attachment)
