from odoo import models, fields, api, _
from odoo.exceptions import UserError
import xlrd
import base64
import logging
from datetime import datetime, timedelta

_logger = logging.getLogger(__name__)

class NvoccImportFormA(models.TransientModel):
    _name = 'nvocc.import.form.a'
    _description = 'Wizard Import Form A'

    file_data = fields.Binary(string='File Excel')
    file_name = fields.Char(string='Nama File')

    def action_download_template(self):
        return {
            'type': 'ir.actions.act_url',
            'name': 'Contoh Form A',
            'url': '/nvocc/static/contoh_Form_A.xlsx',
            'target': 'new',
        }

    def action_import(self):
        if not self.file_data:
            raise UserError("Upload file dulu bro!")

        try:
            file_content = base64.b64decode(self.file_data)
            book = xlrd.open_workbook(file_contents=file_content)
        except Exception as e:
            raise UserError("Gagal membaca file Excel: %s" % str(e))

        try:
            sheet_ref = book.sheet_by_name('REF')
            self._process_reference(sheet_ref)
        except: pass

        freight_fob_map = {}

        try:
            try: sheet_data = book.sheet_by_name('DATA')
            except: sheet_data = book.sheet_by_index(0)
            master_rec = self._process_header(sheet_data, book.datemode)
            self._process_rows(sheet_data, master_rec, book.datemode, freight_fob_map)
        except Exception as e:
            raise UserError("Error Sheet DATA: %s" % str(e))

        try:
            sheet_barang = book.sheet_by_name('BARANG')
            self._process_barang(sheet_barang, master_rec, freight_fob_map)
        except Exception as e:
            pass

        return {
            'type': 'ir.actions.act_window',
            'name': 'Data Manifest',
            'res_model': 'nvocc.master',
            'res_id': master_rec.id,
            'view_mode': 'form',
            'target': 'current',
        }

    # --- HELPERS ---
    def _clean_str(self, val):
        if val is None: return ""
        if isinstance(val, float): 
            if val.is_integer(): return str(int(val))
            return str(val)
        return str(val).replace('_x000D_', ' ').replace('\r\n', ' ').replace('\n', ' ').strip()
        
    def _clean_uraian_barang(self, val):
        if val is None: return ""
        if isinstance(val, float): 
            if val.is_integer(): return str(int(val))
            return str(val)
        return str(val).replace('_x000D_', '\n').replace('\r', '').strip()

    def _get_val(self, sheet, row, col):
        try:
            val = sheet.cell(row, col).value
            return self._clean_str(val)
        except: return ""

    def _get_int(self, sheet, row, col):
        try:
            val = sheet.cell(row, col).value
            if not val: return 0
            if isinstance(val, str):
                val = val.replace(',', '').replace('.', '').strip()
                if not val.isnumeric(): return 0
            return int(float(val))
        except: return 0

    def _get_float(self, sheet, row, col):
        try:
            val = sheet.cell(row, col).value
            if not val: return 0.0
            if isinstance(val, str): val = val.replace(',', '.')
            return float(val)
        except: return 0.0

    def _get_date(self, sheet, row, col, datemode=0):
        try:
            cell = sheet.cell(row, col)
            val = cell.value
            if cell.ctype == 3 or (cell.ctype == 2 and isinstance(val, float) and val > 1000): 
                y, m, d, h, i, s = xlrd.xldate_as_tuple(val, datemode)
                return "{}-{:02d}-{:02d}".format(y, m, d)
            return False
        except: return False

    def _get_datetime(self, sheet, row, col, datemode=0):
        try:
            cell = sheet.cell(row, col)
            val = cell.value
            if cell.ctype == 3 or (cell.ctype == 2 and isinstance(val, float) and val > 1000): 
                y, m, d, h, i, s = xlrd.xldate_as_tuple(val, datemode)
                dt_obj = datetime(y, m, d, h, i, s)
                dt_utc = dt_obj - timedelta(hours=7)
                return dt_utc.strftime('%Y-%m-%d %H:%M:%S')
            return False
        except Exception as e:
            _logger.error("Error datetime convert: %s", str(e))
            return False

    def _get_ref_id(self, code, master_type):
        if not code: return False
        clean = str(code).strip()
        ref = self.env['nvocc.reference'].search([('name', '=', clean), ('kode_master', '=', master_type)], limit=1)
        if not ref:
            ref = self.env['nvocc.reference'].search([('name', '=ilike', clean), ('kode_master', '=', master_type)], limit=1)
        return ref.id if ref else False

    def _process_reference(self, sheet):
        for r in range(2, sheet.nrows):
            try:
                kd, ur = self._get_val(sheet, r, 0), self._get_val(sheet, r, 1)
                if kd: self._create_ref(kd, ur, 1) 
            except: pass
            try:
                kd, ur = self._get_val(sheet, r, 3), self._get_val(sheet, r, 4)
                if kd: self._create_ref(kd, ur, 2) 
            except: pass
            try:
                kd, ur = self._get_val(sheet, r, 6), self._get_val(sheet, r, 7)
                if kd: self._create_ref(kd, ur, 3) 
            except: pass
            try:
                kd, ur = self._get_val(sheet, r, 9), self._get_val(sheet, r, 10)
                if kd: self._create_ref(kd, ur, 4) 
            except: pass
            try:
                kd, ur = self._get_val(sheet, r, 12), self._get_val(sheet, r, 13)
                if kd: self._create_ref(kd, ur, 5) 
            except: pass
            try:
                kd, ur = self._get_val(sheet, r, 14), self._get_val(sheet, r, 15)
                if kd: self._create_ref(kd, ur, 6) 
            except: pass
            try:
                kd_identitas, ur_identitas = self._get_val(sheet, r, 17), self._get_val(sheet, r, 18)
                if kd_identitas: self._create_ref(kd_identitas, ur_identitas, 8) 
            except: pass

    def _create_ref(self, kode, uraian, master_id):
        existing = self.env['nvocc.reference'].search([('name', '=', kode), ('kode_master', '=', master_id)], limit=1)
        if not existing:
            self.env['nvocc.reference'].create({'name': kode, 'uraian': uraian, 'kode_master': master_id})

    def _process_header(self, sheet, datemode):
        no_master = self._get_val(sheet, 0, 1) 
        if not no_master: no_master = "DRAFT-" + datetime.now().strftime('%Y%m%d%H%M')
        tgl_master = self._get_date(sheet, 1, 1, datemode) or fields.Date.today()
        
        company = self.env.user.company_id
        alamat_parts = [
            company.street2 or '',
            company.city or '',
            company.country_id.name if company.country_id else ''
        ]
        alamat_lengkap = " ".join([part for part in alamat_parts if part])

        tgl_berangkat_excel = self._get_datetime(sheet, 2, 5, datemode) 
        tgl_tiba_excel = self._get_datetime(sheet, 3, 5, datemode)

        default_kantor_id = self._get_ref_id('060100', 7)
        
        vals = {
            'name': no_master,
            'tanggalBl': tgl_master,
            'kodeKantor': default_kantor_id,
            'namaPerusahaan': company.name or '',
            'idPerusahaan' : company.vat or '',
            'alamatPerusahaan' : alamat_lengkap,
            'tanggalBerangkat': tgl_berangkat_excel,
            'tanggalTiba': tgl_tiba_excel,
            'kodeNegara': self._get_ref_id(self._get_val(sheet, 3, 1), 1),
            'kodePelabuhanAsal': self._get_ref_id(self._get_val(sheet, 4, 1), 2),
            'kodePelabuhanTransit': self._get_ref_id(self._get_val(sheet, 5, 1), 2),
            'kodePelabuhanBongkar': self._get_ref_id(self._get_val(sheet, 6, 1), 2),
            'nomorContainer': self._get_val(sheet, 8, 1),
            'jenisContainer': self._get_ref_id(self._get_val(sheet, 9, 1), 4),
            'ukuranContainer': self._get_ref_id(self._get_val(sheet, 10, 1), 3),
            'nomorSegel': self._get_val(sheet, 11, 1),
            'namaSaranaPengangkut': self._get_val(sheet, 12, 1),
            'nomorVoyage': self._get_val(sheet, 4, 5),   
            'imoNumber': self._get_val(sheet, 5, 5),     
            'callSign': self._get_val(sheet, 13, 1),     
        }
        
        existing = self.env['nvocc.master'].search([('name', '=', no_master)], limit=1)
        if existing:
            existing.write(vals)
            existing.dataBls.unlink()
            return existing
        else:
            return self.env['nvocc.master'].create(vals)

    def _process_rows(self, sheet, master_rec, datemode, freight_fob_map):
        start_row = 14
        for r in range(start_row, sheet.nrows):
            no_pos = self._get_val(sheet, r, 0)
            if not no_pos or str(no_pos).upper() in ['NO POS', 'NO SUB POS']: continue
            col2 = self._get_val(sheet, r, 1)
            if str(no_pos) in ['1', '2', '3'] and str(col2) in ['2', '3', '4']: continue

            clean_sub = self._clean_str(sheet.cell(r, 0).value) 

            freight_val = self._get_float(sheet, r, 20) 
            fob_val = self._get_float(sheet, r, 21)     
            
            if clean_sub:
                freight_fob_map[clean_sub] = {
                    'freight': freight_val,
                    'fob': fob_val
                }

            vals = {
                'masterBlId': master_rec.id,
                'nomorPos': no_pos,
                'nomorSubPos': clean_sub, 
                'nomorHostBl': self._get_val(sheet, r, 1), 
                'tanggalHostBl': self._get_date(sheet, r, 2, datemode), 
                'npwpPenerima': self._get_val(sheet, r, 3), 
                'jenis_id_penerima': self._get_ref_id(self._get_val(sheet, r, 32), 8),
                'namaPenerima': self._get_val(sheet, r, 4),   
                'alamatPenerima': self._get_val(sheet, r, 5),
                'negaraPenerima': self._get_ref_id(self._get_val(sheet, r, 8), 1),
                'id_shipper': self._get_val(sheet, r, 9),
                'jenis_id_pengirim': self._get_ref_id(self._get_val(sheet, r, 33), 8),
                'namaPengirim': self._get_val(sheet, r, 10),    
                'alamatPengirim': self._get_val(sheet, r, 11), 
                'negaraPengirim': self._get_ref_id(self._get_val(sheet, r, 12), 1),
                'namaNotify': "SAME AS CONSIGNEE",
                'alamatNotify': "SAME AS CONSIGNEE",
                'jumlahKemasan': self._get_int(sheet, r, 24),
                'jenisKemasan': self._get_val(sheet, r, 25),
                'berat': self._get_float(sheet, r, 13), 
                'netto': self._get_float(sheet, r, 14), 
                'dimensi': self._get_float(sheet, r, 37),
                'telp_penerima': self._clean_str(sheet.cell(r, 15).value),
                'telp_pengirim': self._clean_str(sheet.cell(r, 16).value), 
                'no_invoice': self._clean_str(sheet.cell(r, 27).value), 
                'tgl_invoice': self._get_date(sheet, r, 28, datemode),
                
                # --- FIELD BARU UNTUK FORM B ---
                'jenis_aju': self._get_val(sheet, r, 29),
                'jenis_pibk': self._get_val(sheet, r, 30),
                'no_sub_sub_pos': '0000',
                'kategori_barang': self._get_val(sheet, r, 34),
            }
            self.env['nvocc.house'].create(vals)

    def _process_barang(self, sheet, master_rec, freight_fob_map):
        for r in range(1, sheet.nrows):
            raw_sub = sheet.cell(r, 0).value
            clean_sub = self._clean_str(raw_sub) 
            if not clean_sub or str(clean_sub).upper() in ['NO SUB POS']: continue
            
            house = self.env['nvocc.house'].search([('masterBlId', '=', master_rec.id), ('nomorSubPos', '=', clean_sub)], limit=1)
            if not house and clean_sub.isdigit():
                house = self.env['nvocc.house'].search([('masterBlId', '=', master_rec.id), ('nomorSubPos', '=', clean_sub.zfill(4))], limit=1)
            if not house and clean_sub.isdigit():
                house = self.env['nvocc.house'].search([('masterBlId', '=', master_rec.id), ('nomorSubPos', '=', str(int(clean_sub)))], limit=1)

            if house:
                raw_uraian = sheet.cell(r, 3).value
                clean_uraian = self._clean_uraian_barang(raw_uraian)
                
                f_data = freight_fob_map.get(clean_sub, {'freight': 0.0, 'fob': 0.0})
                if f_data['freight'] == 0.0 and f_data['fob'] == 0.0 and clean_sub.isdigit():
                    f_data = freight_fob_map.get(clean_sub.zfill(4), {'freight': 0.0, 'fob': 0.0})
                if f_data['freight'] == 0.0 and f_data['fob'] == 0.0 and clean_sub.isdigit():
                    f_data = freight_fob_map.get(str(int(clean_sub)), {'freight': 0.0, 'fob': 0.0})

                self.env['nvocc.goods'].create({
                    'houseId': house.id,
                    'kodeHs': self._get_val(sheet, r, 2),
                    'uraianBarang': clean_uraian,
                    'cif': self._get_float(sheet, r, 4),
                    'freight': f_data['freight'],
                    'fob': f_data['fob'],
                })