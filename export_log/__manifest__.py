# -*- coding: utf-8 -*-
{
    'name': 'Export Log ',
    'summary': 'Log Export',
    'description': """Exporting log values from model. Working only in odoo 11""",
    'category': 'Tools',
    'author': 'Niazi Mahrab',
    'maintainer': 'Niazi Mahrab',
    'version': '11.0.0.1',
    'website': 'https://github.com/niazi12/Export-Log/tree/11.0/export_log',
    'license': 'LGPL-3',
    'depends': ['base'],
    'images': ['static/description/banner.gif'],

    'data': [
        'security/ir.model.access.csv',
        'wizards/export_log_wizard_view.xml',
        'views/export_log_menu_views.xml',
        'views/excel_repot_out_views.xml',

    ],

    'installable': True,
    'application': True,
    'auto_install': False,
}
