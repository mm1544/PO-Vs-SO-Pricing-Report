from odoo import fields, models


class ResConfigSettings(models.TransientModel):
    _inherit = 'res.config.settings'
    # config_parameter='licence_expiration_report.recipient_email' will create a system parameter.
    recipient_email = fields.Char(string='Recipient Email', config_parameter='po_vs_so_pricing_report.recipient_email',
                                  help='Can be added multiple, comma separated email addresses')

    sender_email = fields.Char(
        string='Sender Email', config_parameter='po_vs_so_pricing_report.sender_email')

    cc_email = fields.Char(
        string='CC Email', config_parameter='po_vs_so_pricing_report.cc_email')

    reply_to_email = fields.Char(
        string='Reply To Email', config_parameter='po_vs_so_pricing_report.reply_to_email')

    email_company_name = fields.Char(string='Email Company Name', config_parameter='po_vs_so_pricing_report.email_company_name',
                                     help='Company name added to the email')
