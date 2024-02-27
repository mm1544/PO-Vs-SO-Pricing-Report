{
    'name': 'PO vs SO Pricing Report',
    'version': '1.0',
    'category': 'Generic Modules/Others',
    'summary': 'Generates report for products where POLine Unit Price is lower than SOLine Cost and sends to the designated person.',
    'sequence': '1',
    'author': 'Martynas Minskis',
    'depends': ['sale'],
    'demo': [],
    'data': [

        # Sequence: security, data, wizards, views
        'views/pricing_report.xml',
        'views/res_config_settings_views.xml',

    ],
    'demo': [],
    'qweb': [],

    'installable': True,
    'application': True,
    'auto_install': False,
    #     'licence': 'OPL-1',
}