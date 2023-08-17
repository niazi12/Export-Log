from odoo import api, exceptions, fields, models, _
import base64, io
import xlwt
from datetime import datetime
from bs4 import BeautifulSoup


class WitzExportLog(models.TransientModel):
    _name = 'witz.export.log.wizard'

    model_id = fields.Many2one('ir.model', string='Model Name', required=True)

    @api.multi
    def witz_export_action(self):
        """ Generate Export Excel format """
        row = 0
        # Generate File name format
        time_format = datetime.today().strftime("%Y%m%d%H%M%S")
        filename = 'ExportLog-' + time_format + '.xls'

        # Create a workbook and add a worksheet.
        workbook = xlwt.Workbook(encoding="UTF-8")
        worksheet = workbook.add_sheet('Export Log')
        first_col = worksheet.col(0)
        first_col.width = 256 * 15  # 30 characters wide

        data = self.generate_data_all(model=self.model_id.model)
        self._write_in_export_sheet(worksheet, data, row, rec=False, rec_name=False, rec_id=False)
        fp = io.BytesIO()
        workbook.save(fp)
        export_id = self.env['excel.report.out'].create(
            {'excel_file': base64.encodestring(fp.getvalue()), 'file_name': filename})
        fp.close()

        return {
            # 'name': 'Export Log ',
            'name': 'Export Log ',
            'view_mode': 'form',
            'res_id': export_id.id,
            'res_model': 'excel.report.out',
            'view_type': 'form',
            'type': 'ir.actions.act_window',
            'target': 'new',
        }

    def _write_in_export_sheet(self, worksheet, data, row, rec, rec_name=False, rec_id=False):
        # Column Style
        header_style = xlwt.easyxf('pattern: pattern solid, fore_colour dark_purple;'
                                   'font: color white, bold True;')
        header_style1 = xlwt.easyxf('pattern: pattern solid, fore_colour dark_purple;'
                                    'font: color white, bold True;'
                                    'alignment: horizontal center, vertical center;')

        # File name Style
        file_name_style = xlwt.easyxf('pattern: pattern solid, fore_colour yellow;')
        file_name_style1 = xlwt.easyxf('pattern: pattern solid, fore_colour red;')

        # File Header Style
        file_header_style = xlwt.easyxf('font: bold on; align: wrap on, vert centre, horiz center')
        row = 0
        col = 0
        mer_row = row + 1
        if rec_id:
            worksheet.write_merge(row, mer_row, col, col, 'Record Name', header_style1)
            col += 1
            worksheet.write_merge(row, mer_row, col, col, 'Record ID', header_style1)
            col += 1

        worksheet.write_merge(row, mer_row, col, col, 'Old Value', header_style1)
        col += 1
        worksheet.write_merge(row, mer_row, col, col, 'New Value', header_style1)
        col += 1
        worksheet.write_merge(row, mer_row, col, col, 'Responsible', header_style1)
        col += 1
        worksheet.write_merge(row, mer_row, col, col, 'Date and Time', header_style1)

        data_lists = []
        for res in data:
            data_lists.append(res)

        row = 2
        for indx in range(0, len(data_lists)):
            col = 0
            if rec_id:
                worksheet.write(row, col, format(data_lists[indx]["rec_name"]))
                col += 1
                worksheet.write(row, col, format(data_lists[indx]["rec_id"]))
                col += 1
                # col += 1
            worksheet.write(row, col, format(data_lists[indx]['field_name']))
            col += 1
            worksheet.write(row, col, format(data_lists[indx]['new_value']))
            col += 1
            worksheet.write(row, col, format(data_lists[indx]['l']))
            col += 1
            worksheet.write(row, col, format(data_lists[indx]['write_date']))
            col += 1

            row += 1

        return

    def generate_data_all(self, model=None, rec=None):
        # Model Filtering
        condition_str = ""
        if model:
            condition_str = "mm.model = '{}'".format(model)
        if rec:
            if len(rec) > 1:
                condition_str = condition_str + " and mm.res_id in {}".format(tuple(rec))
            else:
                condition_str = condition_str + " and mm.res_id = {}".format(rec[0])

        sql = """
              SELECT  mtv.field_type as subject, mtv.field as field_status, mtv.field_desc as field_desc, 
              mtv.old_value_integer as old_value_integer, mtv.new_value_integer as new_value_integer, 
              mtv.old_value_float as old_value_float, mtv.new_value_float as new_value_float, 
              mtv.old_value_char as old_value_char, mtv.new_value_char as new_value_char, 
              mtv.old_value_text as old_value_text, mtv.new_value_text as new_value_text, 
              mtv.old_value_datetime as old_value_datetime, mtv.new_value_datetime as new_value_datetime,
              mtv.old_value_monetary as old_value_monetary, mtv.new_value_monetary as new_value_monetary,
              mtv.write_date as write_date, 
                usr.id as user_id,
                res.name as responsible,
                mm.res_id as res_id, mm.body as body, 
                mm.write_date as write_date_mail_message,
                mm.write_uid as write_uid_mail_message,
                res.create_uid as write_uid_res_user
              FROM mail_tracking_value as mtv 
                  FULL JOIN mail_message as mm on mtv.mail_message_id = mm.id 
                  LEFT JOIN res_users as usr on usr.id = mm.create_uid
				  LEFT JOIN res_partner as res on res.id = usr.partner_id 
				  LEFT JOIN hr_employee as hr on hr.id = usr.id

				  WHERE  %s 
                    ORDER BY mm.res_id
                      """ % (str(condition_str))
        self.env.cr.execute(sql)
        results = self.env.cr.dictfetchall()
        emt_dic_list = []
        # Adding old and new data
        for res in results:
            soup = BeautifulSoup(res['body'], 'html.parser')
            res['str_body'] = soup.get_text()
            emt_dic = {}
            if res['old_value_char'] and res['old_value_char'] != '':
                old_val = res['old_value_char']
            elif res['old_value_text'] and res['old_value_text'] != '':
                old_val = res['old_value_text']
            elif res['old_value_monetary'] and res['old_value_monetary'] > 0:
                old_val = res['old_value_monetary']
            elif res['old_value_datetime'] and res['old_value_datetime'] != '':
                old_val = res['old_value_datetime']
            elif res['old_value_float'] and res['old_value_float'] > 0:
                old_val = res['old_value_float']
            elif res['old_value_integer']:
                if res['subject'] == 'boolean' and res['old_value_integer'] == 1:
                    old_val = 'True'
                elif res['subject'] == 'boolean' and res['old_value_integer'] == 0:
                    old_val = 'False'
                else:
                    old_val = res['old_value_integer']
            else:
                old_val = ''

            if res['new_value_char'] and res['new_value_char'] != '':
                new_val = res['new_value_char']
            elif res['str_body'] and res['str_body'] != '':
                new_val = res['str_body']
            elif res['new_value_text'] and res['new_value_text'] != '':
                new_val = res['new_value_text']
            elif res['new_value_monetary'] and res['new_value_monetary'] > 0:
                new_val = res['new_value_monetary']
            elif res['new_value_datetime'] and res['new_value_datetime'] != '':
                new_val = res['new_value_datetime']
            elif res['new_value_float'] and res['new_value_float'] > 0:
                new_val = res['new_value_float']
            elif res['new_value_integer']:
                if res['subject'] == 'boolean' and res['new_value_integer'] == 1:
                    new_val = 'True'
                elif res['subject'] == 'boolean' and res['new_value_integer'] == 0:
                    new_val = 'False'
                else:
                    new_val = res['new_value_integer']

            else:
                new_val = ''

            updated_date = datetime.strptime(str(res['write_date']), '%Y-%m-%d %H:%M:%S.%f') if res[
                'write_date'] else datetime.strptime(str(res['write_date_mail_message']), '%Y-%m-%d %H:%M:%S.%f')
            updated_date = updated_date.strftime('%d %b, %Y %H:%M:%S')

            if res['user_id']:
                updated_responsible = self.env['hr.employee'].search([('user_id', '=', res['user_id'])])
                if updated_responsible:
                    res['responsible'] = updated_responsible.name
            if not old_val:
                old_val = 'Empty'
            new_log_value = "{} -> {}".format(old_val, new_val)

            res_id = res['res_id']
            emt_dic.update({'field_name': res['field_desc'],
                            'old_value': old_val, 'new_value': new_log_value, 'l': res['responsible'],
                            'write_date': updated_date, 'res_id': res_id})
            emt_dic_list.append(emt_dic)
        return emt_dic_list
