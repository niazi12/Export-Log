from odoo import fields,models,_

class ExcelReportOut(models.TransientModel):
    _name = 'excel.report.out'
    _description = 'Excel report'

    excel_file = fields.Binary('Excel Report')
    file_name = fields.Char('Excel File', size=64)