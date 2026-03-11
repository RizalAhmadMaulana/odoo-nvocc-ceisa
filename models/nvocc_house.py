from odoo import models, fields, api
from odoo.exceptions import UserError

class NvoccHouse(models.Model):
    _name = 'nvocc.house'
    _description = 'Detail House BL (Pos)'
    _rec_name = 'nomorHostBl'

    masterBlId = fields.Many2one('nvocc.master', string='Master BL', ondelete='cascade', required=True)
    parent_state = fields.Selection(related='masterBlId.state', string='Status Master', store=False) 
    
    nomorPos = fields.Char(string='Nomor Pos', default="000000000000")
    nomorSubPos = fields.Char(string='No Sub Pos') 
    nomorHostBl = fields.Char(string='Nomor Host BL')
    tanggalHostBl = fields.Date(string='Tanggal Host BL')
    
    id_shipper = fields.Char(string='ID Shipper') 
    namaPengirim = fields.Char(string='Nama Pengirim')
    alamatPengirim = fields.Text(string='Alamat Pengirim')
    negaraPengirim = fields.Many2one('nvocc.reference', string='Negara Pengirim', domain=[('kode_master','=',1)])
    jenis_id_pengirim = fields.Many2one('nvocc.reference', string='Jenis ID Pengirim', domain=[('kode_master','=',8)])
    
    namaPenerima = fields.Char(string='Nama Penerima')
    alamatPenerima = fields.Text(string='Alamat Penerima')
    npwpPenerima = fields.Char(string='NPWP/ID Penerima')
    negaraPenerima = fields.Many2one('nvocc.reference', string='Negara Penerima', domain=[('kode_master','=',1)])
    jenis_id_penerima = fields.Many2one('nvocc.reference', string='Jenis ID Penerima', domain=[('kode_master','=',8)])
    
    namaNotify = fields.Char(string='Nama Notify')
    alamatNotify = fields.Text(string='Alamat Notify')
    
    jumlahKemasan = fields.Integer(string='Jumlah Kemasan')
    jenisKemasan = fields.Char(string='Jenis Kemasan')
    
    berat = fields.Float(string='Bruto (KGM)')
    netto = fields.Float(string='Netto (KGM)')
    
    dimensi = fields.Float(string='Dimensi (M3)', digits=(16, 3))
    marking = fields.Char(string='Marking')
    
    blHs = fields.One2many('nvocc.goods', 'houseId', string='Barang (HS)')
    telp_penerima = fields.Char(string='Telp Penerima')
    telp_pengirim = fields.Char(string='Telp Pengirim')
    no_invoice = fields.Char(string='No Invoice')
    tgl_invoice = fields.Date(string='Tgl Invoice')

    jenis_aju = fields.Char(string='Jenis Aju')
    jenis_pibk = fields.Char(string='Jenis PIBK')
    no_sub_sub_pos = fields.Char(string='No Sub Sub Pos', default='0000')
    kategori_barang = fields.Char(string='Kategori Barang')

    @api.model
    def create(self, vals):       
        if 'masterBlId' in vals:
            master = self.env['nvocc.master'].browse(vals['masterBlId'])
            if master.state != 'draft':
                raise UserError("Master Data sudah Validate! Tidak bisa tambah House BL baru.")
        return super(NvoccHouse, self).create(vals)

    @api.multi
    def write(self, vals):
        for rec in self:
            if rec.masterBlId.state != 'draft':
                raise UserError("Data terkunci! Kembalikan status Master ke Draft untuk mengedit.")
        return super(NvoccHouse, self).write(vals)

    @api.multi
    def unlink(self):
        for rec in self:
            if rec.masterBlId.state != 'draft':
                raise UserError("Data terkunci! Tidak bisa menghapus House BL.")
        return super(NvoccHouse, self).unlink()


class NvoccGoods(models.Model):
    _name = 'nvocc.goods'
    _description = 'Detail Barang (HS Code)'

    houseId = fields.Many2one('nvocc.house', string='House BL', ondelete='cascade')
    parent_state = fields.Selection(related='houseId.parent_state', store=False)

    seriHs = fields.Integer(string='Seri HS')
    kodeHs = fields.Char(string='Kode HS') 
    uraianBarang = fields.Text(string='Uraian Barang')
    cif = fields.Float(string='Nilai CIF')
    freight = fields.Float(string='Freight')
    fob = fields.Float(string='FOB')

    @api.multi
    def write(self, vals):
        for rec in self:
            if rec.parent_state != 'draft':
                raise UserError("Data terkunci! Kembalikan status Master ke Draft untuk mengedit.")
        return super(NvoccGoods, self).write(vals)

    @api.multi
    def unlink(self):
        for rec in self:
            if rec.parent_state != 'draft':
                raise UserError("Data terkunci! Tidak bisa menghapus Barang.")
        return super(NvoccGoods, self).unlink()