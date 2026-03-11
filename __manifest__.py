# -*- coding: utf-8 -*-
{
    'name': 'NVOCC Manifest Pecah Pos',
    'version': '12.0.1.0.0',
    'summary': 'Modul NVOCC untuk Pecah Pos Form A ke CEISA',
    'description': 'Modul Management Data Manifest, House BL, dan Barang NVOCC',
    'category': 'Customs',
    'author': 'Kalijaga Dev',
    'website': '-',
    'depends': ['base'],
    'data': [
        'security/ir.model.access.csv',
        
        # 1. Load Views & Actions DULU (Supaya ID-nya terbentuk)
        'views/nvocc_reference_views.xml',
        'views/nvocc_views.xml',  
        'wizards/wiz_import_form_a.xml', # <-- Action wizard ada di sini
        
        # 2. Baru Load Menu TERAKHIR (Supaya bisa panggil ID di atas)
        'views/nvocc_menu.xml',
    ],
    'installable': True,
    'application': True,
}