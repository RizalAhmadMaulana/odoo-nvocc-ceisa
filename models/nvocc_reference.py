from odoo import models, fields, api

class NvoccReference(models.Model):
    _name = 'nvocc.reference'
    _description = 'Referensi Data NVOCC'

    name = fields.Char(string='Kode', required=True)
    uraian = fields.Char(string='Uraian')
    # Kode Master untuk membedakan jenis (1=Negara, 2=Pelabuhan, 3=Kemasan, 4=Satuan)
    kode_master = fields.Integer(string='Kode Master', index=True) 

    def name_get(self):
        result = []
        for rec in self:
            name = rec.name
            if rec.uraian:
                name = '%s - %s' % (name, rec.uraian)
            result.append((rec.id, name))
        return result