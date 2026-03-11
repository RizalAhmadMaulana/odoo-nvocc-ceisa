from odoo import models, fields, api, _
from odoo.exceptions import UserError
import json
import base64
import logging
from datetime import datetime, timedelta
from collections import OrderedDict
import xml.etree.ElementTree as ET
from xml.dom import minidom
import xlwt
import io

_logger = logging.getLogger(__name__)

class NvoccMaster(models.Model):
    _name = 'nvocc.master'
    _description = 'Header Manifest NVOCC'
    _order = 'tanggalBl desc, name desc'

    # --- Header Data ---
    name = fields.Char(string='No Master BL', required=True)
    kodeKantor = fields.Many2one('nvocc.reference', string='Kode Kantor', domain=[('kode_master','=',7)])
    jenisManifes = fields.Char(string='Jenis Manifes', default='11')
    idPerusahaan = fields.Char(string='NPWP Perusahaan')
    namaPerusahaan = fields.Char(string='Nama Perusahaan')
    alamatPerusahaan = fields.Text(string='Alamat Perusahaan')
    
    nomorAju = fields.Char(string='Nomor Aju', readonly=True, copy=False)
    tanggalBl = fields.Date(string='Tgl Master BL', required=True)

    # --- Data Angkut & Waktu ---
    nomorVoyage = fields.Char(string='Nomor Voyage')
    modePengangkut = fields.Selection([('1','Laut'),('2','Udara')], string='Mode Pengangkut', default='1')
    namaSaranaPengangkut = fields.Char(string='Nama Sarana Pengangkut')
    
    imoNumber = fields.Char(string='IMO Number')
    callSign = fields.Char(string='Call Sign') 
    
    kodeNegara = fields.Many2one('nvocc.reference', string='Kode Negara (Asal)', domain=[('kode_master','=',1)])
    tanggalBerangkat = fields.Datetime(string='Tanggal Berangkat')
    tanggalTiba = fields.Datetime(string='Tanggal Tiba')
    
    # --- Pelabuhan ---
    kodePelabuhanAsal = fields.Many2one('nvocc.reference', string='Pelabuhan Asal', domain=[('kode_master','=',2)])
    kodePelabuhanBongkar = fields.Many2one('nvocc.reference', string='Pelabuhan Bongkar', domain=[('kode_master','=',2)])
    kodePelabuhanTransit = fields.Many2one('nvocc.reference', string='Pelabuhan Transit', domain=[('kode_master','=',2)])
    
    # --- Container ---
    nomorContainer = fields.Char(string='Nomor Container')
    jenisContainer = fields.Many2one('nvocc.reference', string='Jenis Container', domain=[('kode_master','=',4)])
    ukuranContainer = fields.Many2one('nvocc.reference', string='Ukuran Container', domain=[('kode_master','=',3)])
    nomorSegel = fields.Char(string='Nomor Segel')

    # --- Data Respon CEISA (Ditambahkan) ---
    nomor_bc11 = fields.Char(string='No BC 1.1')
    tanggal_bc11 = fields.Date(string='Tgl BC 1.1')
    nomor_pos_bc11 = fields.Char(string='No Pos BC 1.1')
    
    # --- Relasi House BL ---
    dataBls = fields.One2many('nvocc.house', 'masterBlId', string='Data BLs (House)')
    jumlahPos = fields.Integer(string='Jumlah Pos', compute='_compute_jumlah_pos')
    
    state = fields.Selection([
        ('draft', 'Draft'),
        ('ready', 'Ready to Send'),
        ('sent', 'Sent to CEISA'),
        ('done', 'Completed')
    ], string='Status', default='draft', track_visibility='onchange')

    # --- FILE JSON & XML ---
    json_file = fields.Binary(string='File JSON CEISA', readonly=True)
    json_filename = fields.Char(string='Nama File JSON')
    
    xml_file = fields.Binary(string='File XML CEISA', readonly=True)
    xml_filename = fields.Char(string='Nama File XML')

    # --- FILE FORM B ---
    form_b_file = fields.Binary(string='File Form B', readonly=True)
    form_b_filename = fields.Char(string='Nama File Form B')

    @api.depends('dataBls')
    def _compute_jumlah_pos(self):
        for rec in self:
            rec.jumlahPos = len(rec.dataBls)

    def action_view_house_bl(self):
        self.ensure_one()
        return {
            'name': _('Data BLs (House)'),
            'type': 'ir.actions.act_window',
            'res_model': 'nvocc.house',
            'view_mode': 'tree,form',
            'domain': [('masterBlId', '=', self.id)],
            'context': {'default_masterBlId': self.id},
        }

    @api.model
    def create(self, vals):
        if not vals.get('nomorAju'):
            npwp_raw = vals.get('idPerusahaan') or self.env.user.company_id.vat or ""
            npwp_clean = "".join([c for c in str(npwp_raw) if c.isdigit()])
            
            # UBAH: Ambil 6 digit pertama NPWP
            npwp_6_digit = npwp_clean[:6].ljust(6, '0')
            
            # UBAH: Format tanggal jadi YYYYMMDD
            today_str = datetime.now().strftime('%Y%m%d')
            
            # Gabungkan menjadi ONEERP + 6 Digit NPWP + YYYYMMDD
            prefix = "ONEERP{}{}".format(npwp_6_digit, today_str)
            
            last_rec = self.env['nvocc.master'].search([('nomorAju', '=ilike', prefix + '%')], order='nomorAju desc', limit=1)
            if last_rec and last_rec.nomorAju:
                try:
                    last_seq = int(last_rec.nomorAju[-6:])
                    new_seq = last_seq + 1
                except ValueError:
                    new_seq = 1
            else:
                new_seq = 1
                
            vals['nomorAju'] = "{}{:06d}".format(prefix, new_seq)
            
        return super(NvoccMaster, self).create(vals)

    def action_generate_json(self):
        for rec in self:
            if not rec.dataBls:
                raise UserError("Data House BL masih kosong!")
            
            if rec.json_file or not rec.nomorAju:
                npwp_raw = rec.idPerusahaan or self.env.user.company_id.vat or ""
                npwp_clean = "".join([c for c in str(npwp_raw) if c.isdigit()])
                
                npwp_6_digit = npwp_clean[:6].ljust(6, '0')
                today_str = datetime.now().strftime('%Y%m%d')
                
                prefix = "ONEERP{}{}".format(npwp_6_digit, today_str)
                
                last_rec = self.env['nvocc.master'].search([('nomorAju', '=ilike', prefix + '%')], order='nomorAju desc', limit=1)
                
                if last_rec and last_rec.nomorAju:
                    try:
                        last_seq = int(last_rec.nomorAju[-6:])
                        new_seq = last_seq + 1
                    except ValueError:
                        new_seq = 1
                else:
                    new_seq = 1
                    
                new_aju = "{}{:06d}".format(prefix, new_seq)
                rec.write({'nomorAju': new_aju})

            def clean_text(val):
                if not val: return " "
                return str(val).replace('_x000D_', ' ').replace('\r\n', ' ').replace('\n', ' ').strip()

            def fmt_date(d): return d.strftime('%Y-%m-%d') if d else ""
            def fmt_time(t): return (t + timedelta(hours=7)).strftime('%Y-%m-%d %H:%M:%S') if t else ""
            def get_ref(r): return clean_text(r.name) if r else ""

            total_items = len(rec.dataBls)
            total_pkg = sum(h.jumlahKemasan for h in rec.dataBls)
            total_brutto = sum(h.berat for h in rec.dataBls)
            total_volume = sum(h.dimensi for h in rec.dataBls)
            total_container = 1 if rec.nomorContainer else 0

            # =========================================================
            # PART 1: GENERATE XML
            # =========================================================
            root_xml = ET.Element("Declaration")
            
            ET.SubElement(root_xml, "FunctionalReferenceID").text = str(rec.id)
            ET.SubElement(root_xml, "FunctionCode").text = "IS"
            ET.SubElement(root_xml, "StatusCode").text = "EM"
            ET.SubElement(root_xml, "ID").text = clean_text(rec.nomorAju)
            
            ET.SubElement(root_xml, "DeclarationOfficeID").text = get_ref(rec.kodeKantor)
            
            btm = ET.SubElement(root_xml, "BorderTransportMeans")
            ET.SubElement(btm, "ArrivalDateTime").text = fmt_time(rec.tanggalTiba)
            ET.SubElement(btm, "DepartureDateTime").text = fmt_time(rec.tanggalBerangkat)
            ET.SubElement(btm, "TypeCode").text = str(rec.modePengangkut) if rec.modePengangkut else "1"
            ET.SubElement(btm, "Name").text = clean_text(rec.namaSaranaPengangkut)
            
            carrier = ET.SubElement(btm, "Carrier")
            ET.SubElement(carrier, "TypeID").text = "1"
            ET.SubElement(carrier, "ID").text = clean_text(rec.idPerusahaan)
            ET.SubElement(carrier, "Name").text = clean_text(rec.namaPerusahaan)
            ET.SubElement(carrier, "Address").text = clean_text(rec.alamatPerusahaan)
            
            master_xml = ET.SubElement(btm, "Master")
            ET.SubElement(master_xml, "Name").text = " "
            
            ET.SubElement(btm, "JurneyID").text = clean_text(rec.nomorVoyage)
            ET.SubElement(btm, "JurneyDate").text = fmt_date(rec.tanggalBerangkat) or fmt_date(rec.tanggalBl)
            ET.SubElement(btm, "CallSign").text = clean_text(rec.callSign)
            
            ET.SubElement(btm, "MmsiID").text = " "
            ET.SubElement(btm, "RegistrationNationalityCode").text = clean_text(rec.imoNumber) 
            ET.SubElement(btm, "CountryCode").text = get_ref(rec.kodeNegara)
            
            ET.SubElement(root_xml, "TotalItem").text = str(total_items)
            ET.SubElement(root_xml, "TotalContainerQuantity").text = str(total_container)
            ET.SubElement(root_xml, "TotalPackageQuantity").text = str(total_pkg)
            ET.SubElement(root_xml, "Brutto").text = "{:.4f}".format(total_brutto)
            ET.SubElement(root_xml, "Volume").text = "{:.3f}".format(total_volume)
            
            msg_name = "V" + clean_text(rec.nomorAju)
            ET.SubElement(root_xml, "MessageName").text = msg_name
            ET.SubElement(root_xml, "MessageID").text = "V"
            
            ET.SubElement(root_xml, "LoadingLocation").text = get_ref(rec.kodePelabuhanAsal)
            ET.SubElement(root_xml, "RoutingRegionIdentificationID").text = " "
            ET.SubElement(root_xml, "TransitLocation").text = get_ref(rec.kodePelabuhanTransit) or get_ref(rec.kodePelabuhanAsal)
            ET.SubElement(root_xml, "FirstArrivalLocationID").text = get_ref(rec.kodePelabuhanBongkar)
            ET.SubElement(root_xml, "UnLoadingLocation").text = get_ref(rec.kodePelabuhanBongkar)
            
            ET.SubElement(root_xml, "UnloadingOfficeLocation").text = get_ref(rec.kodeKantor)
            
            ET.SubElement(root_xml, "SenderName").text = clean_text(rec.namaPerusahaan)
            ET.SubElement(root_xml, "ModuleID").text = msg_name
            ET.SubElement(root_xml, "TaxID").text = clean_text(rec.idPerusahaan)

            consignment = ET.SubElement(root_xml, "Consignment")
            
            for h_idx, house in enumerate(rec.dataBls):
                ci = ET.SubElement(consignment, "ConsignmentItem")
                ET.SubElement(ci, "CategoryCode").text = "01"
                ET.SubElement(ci, "RevisionCode").text = "null"
                
                raw_sub = clean_text(house.nomorSubPos).strip()
                if raw_sub and len(raw_sub) == 12:
                    seq_num_val = raw_sub
                elif raw_sub:
                    seq_num_val = raw_sub.zfill(4) + "00000000"
                else:
                    seq_num_val = str(h_idx + 1).zfill(4) + "00000000"
                
                ET.SubElement(ci, "SequenceNumeric").text = seq_num_val
                
                id_modul_val = clean_text(house.nomorHostBl) if clean_text(house.nomorHostBl) else " "
                ET.SubElement(ci, "IDModul").text = id_modul_val
                
                prev_docs = ET.SubElement(ci, "PreviousDocuments")
                
                doc_master = ET.SubElement(prev_docs, "DocumentItem")
                ET.SubElement(doc_master, "TypeCode").text = "704"
                ET.SubElement(doc_master, "ID").text = clean_text(rec.name)
                ET.SubElement(doc_master, "IssueDate").text = fmt_date(rec.tanggalBl)
                
                if house.nomorHostBl:
                    doc_bc = ET.SubElement(prev_docs, "DocumentItem")
                    ET.SubElement(doc_bc, "TypeCode").text = "705"
                    ET.SubElement(doc_bc, "ID").text = clean_text(house.nomorHostBl)
                    ET.SubElement(doc_bc, "IssueDate").text = fmt_date(house.tanggalHostBl) or fmt_date(rec.tanggalBl)

                ET.SubElement(ci, "StatusDocument").text = " "
                ET.SubElement(ci, "LabelDescription").text = "N/M"

                consignor = ET.SubElement(ci, "Consignor")
                ET.SubElement(consignor, "Name").text = clean_text(house.namaPengirim)
                ET.SubElement(consignor, "Address").text = clean_text(house.alamatPengirim)
                
                id_ship_clean = clean_text(house.id_shipper).strip()
                jenis_pengirim = clean_text(house.jenis_id_pengirim.uraian).strip().upper() if house.jenis_id_pengirim else ""
                
                if id_ship_clean and jenis_pengirim:
                    ET.SubElement(consignor, "TaxID").text = "{}-{}".format(jenis_pengirim, id_ship_clean)
                elif id_ship_clean:
                    ET.SubElement(consignor, "TaxID").text = id_ship_clean
                else:
                    ET.SubElement(consignor, "TaxID").text = ""
                    
                ET.SubElement(consignor, "CountryCode").text = get_ref(rec.kodeNegara)

                consignee = ET.SubElement(ci, "Consignee")
                npwp_clean = clean_text(house.npwpPenerima).strip()
                if not npwp_clean:
                    npwp_clean = "0000000000000000"
                
                jenis_penerima = clean_text(house.jenis_id_penerima.uraian).strip().upper() if house.jenis_id_penerima else ""
                if jenis_penerima:
                    ET.SubElement(consignee, "TaxID").text = "{}-{}".format(jenis_penerima, npwp_clean)
                else:
                    ET.SubElement(consignee, "TaxID").text = npwp_clean
                    
                ET.SubElement(consignee, "Name").text = clean_text(house.namaPenerima)
                ET.SubElement(consignee, "Address").text = clean_text(house.alamatPenerima)
                ET.SubElement(consignee, "CountryCode").text = "ID"

                notify = ET.SubElement(ci, "NotifyParty")
                ET.SubElement(notify, "Name").text = "SAME AS CONSIGNEE"
                ET.SubElement(notify, "Address").text = "SAME AS CONSIGNEE"
                ET.SubElement(notify, "CountryCode").text = "ID"

                w_val = "{:.4f}".format(house.berat or 0.0)
                n_val = "{:.4f}".format(house.netto or 0.0)
                v_val = "{:.3f}".format(house.dimensi or 0.0)
                
                ET.SubElement(ci, "GrossMassMeasure").text = w_val
                ET.SubElement(ci, "NetVolumeMeasure").text = v_val
                ET.SubElement(ci, "NetNetWeightMeasure").text = n_val
                
                ET.SubElement(ci, "TotalContainer").text = "1"

                pack = ET.SubElement(ci, "Packaging")
                pkg_code = clean_text(house.jenisKemasan) or "PK"
                ET.SubElement(pack, "ID").text = pkg_code 
                ET.SubElement(pack, "TypeCode").text = pkg_code
                ET.SubElement(pack, "QuantityQuantity").text = str(house.jumlahKemasan or 0)

                ET.SubElement(ci, "ArrivalTransportMeansName").text = clean_text(rec.namaSaranaPengangkut)
                ET.SubElement(ci, "FirstArrivalLocationID").text = get_ref(rec.kodePelabuhanAsal)
                ET.SubElement(ci, "LoadingLocation").text = get_ref(rec.kodePelabuhanAsal)
                ET.SubElement(ci, "UnLoadingLocation").text = get_ref(rec.kodePelabuhanBongkar)
                ET.SubElement(ci, "DocumentDeclaration")
                ET.SubElement(ci, "SequenceNumericCount").text = "0"
                ET.SubElement(ci, "TransitLocation").text = get_ref(rec.kodePelabuhanTransit) or get_ref(rec.kodePelabuhanAsal)
                ET.SubElement(ci, "RoutingRegionIdentificationID").text = get_ref(rec.kodePelabuhanBongkar)

                exchange = ET.SubElement(ci, "Exchange")
                ET.SubElement(exchange, "TotalContainer").text = "1"
                ET.SubElement(exchange, "QuantityQuantity").text = str(house.jumlahKemasan or 0)
                ET.SubElement(exchange, "TypeCode").text = "CT"

                mst = ET.SubElement(ci, "Mst")
                ET.SubElement(mst, "TotalContainer").text = "1"
                ET.SubElement(mst, "QuantityQuantity").text = str(house.jumlahKemasan or 0)
                ET.SubElement(mst, "TypeCode").text = "CT"

                if rec.nomorContainer:
                    te = ET.SubElement(ci, "TransportEquipment")
                    ci_container = ET.SubElement(te, "ContainerItem")
                    ET.SubElement(ci_container, "SequenceNumeric").text = "1"
                    ET.SubElement(ci_container, "ID").text = clean_text(rec.nomorContainer)
                    
                    fcl_code = "FCL" 
                    if get_ref(rec.jenisContainer) == "L": fcl_code = "LCL"
                    elif get_ref(rec.jenisContainer) == "E": fcl_code = "MTY"
                    ET.SubElement(ci_container, "FullnessCode").text = fcl_code
                    
                    ET.SubElement(ci_container, "TypeCode").text = "1"
                    ET.SubElement(ci_container, "CharacteristicCode").text = get_ref(rec.ukuranContainer)
                    ET.SubElement(ci_container, "StatusCode").text = "01" 
                    ET.SubElement(ci_container, "SealID").text = clean_text(rec.nomorSegel)

                commodity = ET.SubElement(ci, "Commodity")
                seq_num = 1 
                for goods in house.blHs:
                    raw_uraian = str(goods.uraianBarang or "")
                    uraian_raw_list = raw_uraian.replace('_x000D_', '\n').replace('\r\n', '\n').replace('\r', '\n').split('\n')
                    uraian_parts = [item.strip() for item in uraian_raw_list if item.strip()]
                    uraian_gabung = ", ".join(uraian_parts)
                    
                    if not uraian_gabung:
                        uraian_gabung = " "
                    
                    uraian_xml = uraian_gabung[:40]
                    cls_xml = ET.SubElement(commodity, "Classification")
                    ET.SubElement(cls_xml, "SequenceNumeric").text = str(seq_num)
                    ET.SubElement(cls_xml, "IdentifierTypeCode").text = "HS"
                    ET.SubElement(cls_xml, "ID").text = clean_text(goods.kodeHs)
                    ET.SubElement(cls_xml, "NameCode").text = uraian_xml
                    
                    seq_num += 1

            # =========================================================
            # PART 2: GENERATE JSON (MENGGUNAKAN ORDERED DICT)
            # =========================================================
            mode_json = dict(rec._fields['modePengangkut'].selection).get(rec.modePengangkut) or "Laut"
            mode_json = mode_json.upper()

            raw_aju = "".join(filter(str.isdigit, rec.nomorAju or ""))
            clean_aju = raw_aju[:26].ljust(26, '0')

            def get_null_str(val):
                return str(val).strip() if val else None

            def get_null_time(t):
                return (t + timedelta(hours=7)).strftime('%Y-%m-%d %H:%M:%S') if t else None

            def get_null_ref(r):
                return clean_text(r.name) if r else None

            # MENGGUNAKAN ORDERED DICT AGAR URUTAN JSON TIDAK ACAK-ACAKAN
            payload = OrderedDict([
                ("kodeKantor", get_ref(rec.kodeKantor)),
                ("jenisManifes", clean_text(rec.jenisManifes) or "11"),
                ("idPerusahaan", "".join(filter(str.isdigit, rec.idPerusahaan or ""))[:15]),
                ("namaPerusahaan", (rec.namaPerusahaan or "")[:100]),
                ("alamatPerusahaan", (rec.alamatPerusahaan or "ALAMAT PERUSAHAAN")[:200]),
                ("nomorAju", clean_aju),
                ("nomorVoyage", clean_text(rec.nomorVoyage)),
                ("tanggalBerangkat", get_null_time(rec.tanggalBerangkat)),
                ("tanggalTiba", get_null_time(rec.tanggalTiba)),
                ("modePengangkut", mode_json),
                ("namaSaranaPengangkut", clean_text(rec.namaSaranaPengangkut)),
                ("imoNumber", clean_text(rec.imoNumber)),
                ("kodeNegara", get_null_ref(rec.kodeNegara)),
                ("asalData", "S"),
                ("flwaktuTempuh", None),
                ("dataKelompokPos", [
                    OrderedDict([
                        ("kodeKelompokPos", "01"),
                        ("jumlah", total_items)
                    ])
                ]),
                ("lampirans", []),
                ("masterBls", [
                    OrderedDict([
                        ("kelompokPos", "01"),
                        ("masterBl", clean_text(rec.name)),
                        ("tanggalBl", fmt_date(rec.tanggalBl)),
                        ("jumlahPos", total_items)
                    ])
                ]),
                ("dataBls", [])
            ])

            for h_idx_json, house in enumerate(rec.dataBls):
                list_petikemas = []
                if rec.nomorContainer:
                    list_petikemas.append(OrderedDict([
                        ("seriContainer", 1),
                        ("nomorContainer", clean_text(rec.nomorContainer)),
                        ("flB3", ""),
                        ("flTerangkut", "01"),
                        ("status", ""),
                        ("typeContainer", "1"),
                        ("ukuranContainer", get_ref(rec.ukuranContainer) or "40"),
                        ("nomorSegel", clean_text(rec.nomorSegel)),
                        ("nomorPolisi", None),
                        ("jenisContainer", get_ref(rec.jenisContainer) or "8"),
                        ("jenisMuat", ""),
                        ("driver", None)
                    ]))

                list_hs = []
                seq_hs_json = 1
                for goods in house.blHs:
                    raw_uraian = str(goods.uraianBarang or "")
                    uraian_raw_list = raw_uraian.replace('_x000D_', '\n').replace('\r\n', '\n').replace('\r', '\n').split('\n')
                    uraian_parts = [item.strip() for item in uraian_raw_list if item.strip()]
                    uraian_json = "\n".join(uraian_parts) if uraian_parts else " "
                    
                    list_hs.append(OrderedDict([
                        ("seri", seq_hs_json),
                        ("kodeHs", clean_text(goods.kodeHs)),
                        ("uraianBarang", uraian_json)
                    ]))
                    seq_hs_json += 1

                raw_sub_json = clean_text(house.nomorSubPos).strip()
                if raw_sub_json and len(raw_sub_json) == 12: seq_num_json = raw_sub_json
                elif raw_sub_json: seq_num_json = raw_sub_json.zfill(4) + "00000000"
                else: seq_num_json = str(h_idx_json + 1).zfill(4) + "00000000"

                house_data = OrderedDict([
                    ("kodeKelompokPos", "01"),
                    ("nomorPos", seq_num_json),
                    ("nomorBl", clean_text(rec.name)),
                    ("tanggalBl", fmt_date(rec.tanggalBl)),
                    ("nomorHostBl", clean_text(house.nomorHostBl)),
                    ("tanggalHostBl", fmt_date(house.tanggalHostBl) or fmt_date(rec.tanggalBl)),
                    ("marking", clean_text(house.marking) if house.marking else "-"),
                    ("npwpPengirim", "".join(filter(str.isdigit, house.id_shipper or ""))),
                    ("namaPengirim", clean_text(house.namaPengirim)),
                    ("alamatPengirim", clean_text(house.alamatPengirim)),
                    ("npwpPenerima", "".join(filter(str.isdigit, house.npwpPenerima or ""))),
                    ("namaPenerima", clean_text(house.namaPenerima)),
                    ("alamatPenerima", clean_text(house.alamatPenerima)),
                    ("npwpNotify", ""),
                    ("namaNotify", clean_text(house.namaNotify) or "SAME AS CONSIGNEE"),
                    ("alamatNotify", clean_text(house.alamatNotify) or "SAME AS CONSIGNEE"),
                    ("motherVessel", None),
                    ("kodePelabuhanAsal", get_ref(rec.kodePelabuhanAsal)),
                    ("kodePelabuhanBongkar", get_ref(rec.kodePelabuhanBongkar)),
                    ("kodePelabuhanTransit", get_ref(rec.kodePelabuhanTransit) or get_ref(rec.kodePelabuhanAsal)),
                    ("kodePelabuhanAkhir", get_ref(rec.kodePelabuhanBongkar)),
                    ("flagKonsolidasi", "N"),
                    ("flagParsial", "N"),
                    ("flagPartof", "N"),
                    ("berat", float("{:.2f}".format(house.berat or 0.0))),
                    ("dimensi", float("{:.3f}".format(house.dimensi or 0.0))),
                    ("jumlahContainer", int(1 if rec.nomorContainer else 0)),
                    ("jenisKemasan", clean_text(house.jenisKemasan) or "CT"),
                    ("jumlahKemasan", int(house.jumlahKemasan or 0)),
                    ("jumlahContainerTertinggal", 0),
                    ("jumlahKemasanTertinggal", 0),
                    ("jumlahKemasanTerangkut", int(house.jumlahKemasan or 0)),
                    ("blHs", list_hs),
                    ("blPetikemasTerangkut", list_petikemas),
                    ("blDokumen", [])
                ])
                
                payload["dataBls"].append(house_data)

            # =========================================================
            # FINISHING: SIMPAN KEDUANYA
            # =========================================================
            xml_raw = ET.tostring(root_xml, encoding='utf-8')
            xml_pretty = minidom.parseString(xml_raw).toprettyxml(indent="    ")
            json_str = json.dumps(payload, indent=4)
            
            base_filename = clean_text(rec.nomorAju) or clean_text(rec.name)
            
            rec.write({
                'json_file': base64.b64encode(json_str.encode('utf-8')),
                'json_filename': "CEISA_NVOCC_{}.json".format(base_filename),
                'xml_file': base64.b64encode(xml_pretty.encode('utf-8')),
                'xml_filename': "CEISA_NVOCC_{}.xml".format(base_filename),
                'state': 'ready'
            })

    def action_generate_form_b(self):
        for rec in self:
            workbook = xlwt.Workbook(encoding='utf-8')
            
            # --- HELPER FORMATTING ---
            def get_str(val): return str(val) if val else ""
            def get_date_str(val):
                if not val: return ""
                val_str = str(val)
                if len(val_str) >= 10:
                    try:
                        dt = datetime.strptime(val_str[:10], '%Y-%m-%d')
                        return dt.strftime('%d-%m-%Y')
                    except:
                        return val_str[:10]
                return val_str

            def get_time_str(val):
                if not val: return ""
                val_str = str(val)
                if len(val_str) >= 19:
                    try:
                        dt = datetime.strptime(val_str[:19], '%Y-%m-%d %H:%M:%S')
                        return dt.strftime('%H:%M:%S')
                    except:
                        return val_str[11:19]
                return "00:00:00"
            
            def get_ref_name(ref_obj):
                return ref_obj.name if ref_obj else ""

            tot_kemasan = sum(h.jumlahKemasan for h in rec.dataBls)
            tot_bruto = sum(h.berat for h in rec.dataBls)
            tot_volume = sum(h.dimensi for h in rec.dataBls)
            tot_container = 1 if rec.nomorContainer else 0

            # ==========================================
            # SHEET 1: KONTAINER
            # ==========================================
            ws_kontainer = workbook.add_sheet('Kontainer')
            cols_kontainer = ['SERI KONTAINER', 'NOMOR KONTAINER', 'UKURAN KONTAINER', 'TIPE KONTAINER', 'JENIS KONTAINER', 'NOMOR SEGEL', 'STATUS KONTAINER']
            for col_idx, col_name in enumerate(cols_kontainer):
                ws_kontainer.write(0, col_idx, col_name)
            
            if rec.nomorContainer:
                row_k = [1, get_str(rec.nomorContainer), get_ref_name(rec.ukuranContainer), "1", get_ref_name(rec.jenisContainer), get_str(rec.nomorSegel), "01"]
                for col_idx, val in enumerate(row_k):
                    ws_kontainer.write(1, col_idx, val)

            # ==========================================
            # SHEET 2: BARANG
            # ==========================================
            ws_barang = workbook.add_sheet('Barang')
            cols_barang = ['NO HOST BLAWB', 'SERI BARANG', 'HS CODE', 'URAIAN BARANG', 'CIF', 'FREIGHT', 'FOB', 'ASURANSI', 'NO SKEP', 'TGL SKEP', 'IMEI1', 'IMEI2', 'BM', 'PPH', 'PPN', 'PPNBM', 'BMTP']
            for col_idx, col_name in enumerate(cols_barang):
                ws_barang.write(0, col_idx, col_name)

            b_row = 1
            for house in rec.dataBls:
                for hs in house.blHs:
                    row_b = [
                        get_str(house.nomorHostBl), hs.seriHs or 1, get_str(hs.kodeHs), get_str(hs.uraianBarang),
                        hs.cif or 0.0, 
                        hs.freight or 0.0,
                        hs.fob or 0.0, 
                        0, "", "", "", "", 0, 0, 0, 0, 0
                    ]
                    for col_idx, val in enumerate(row_b):
                        ws_barang.write(b_row, col_idx, val)
                    b_row += 1

            # ==========================================
            # SHEET 3: DETIL
            # ==========================================
            ws_detil = workbook.add_sheet('Detil')
            cols_detil = ['ID MASTER', 'NOMOR AJU', 'KD KELOMPOK POS', 'NO POS', 'NO SUB POS', 'NO MASTER BLAWB', 'TGL MASTER BLAWB', 'NO HOST BLAWB', 'TGL HOST BLAWB', 'MOTHER VESSEL', 'NPWP CONSIGNEE', 'NAMA CONSIGNEE', 'ALMT CONSIGNEE', 'NEG CONSIGNEE', 'NPWP SHIPPER', 'NAMA SHIPPER', 'ALMT SHIPPER', 'NEG SHIPPER', 'NAMA NOTIFY', 'ALMT NOTIFY', 'NEG NOTIFY', 'PELABUHAN ASAL', 'PELABUHAN TRANSIT', 'PELABUHAN BONGKAR', 'PELABUHAN AKHIR', 'JUMLAH KEMASAN', 'JENIS KEMASAN', 'MERK KEMASAN', 'JUMLAH KONTAINER', 'BRUTO', 'VOLUME', 'FL PARTIAL', 'TOTAL KEMASAN', 'TOTAL KONTAINER', 'STATUS DETIL', 'JENIS ID SHIPPER', 'JENIS ID CONSIGNEE', 'TELP PENERIMA', 'TELP PENGIRIM', 'NO BARCODE', 'NO INVOICE', 'TGL INVOICE', 'JENIS AJU', 'JENIS PIBK', 'NO SUB SUB POS', 'KATEGORI BARANG', 'ID PPMSE', 'NAMA PPMSE']
            for col_idx, col_name in enumerate(cols_detil):
                ws_detil.write(0, col_idx, col_name)

            for idx, house in enumerate(rec.dataBls):
                no_pos_bc = get_str(rec.nomor_pos_bc11) if rec.nomor_pos_bc11 else get_str(house.nomorPos)
                row_d = [
                    str(rec.id), get_str(rec.nomorAju), "01", no_pos_bc, get_str(house.nomorSubPos),
                    get_str(rec.name), get_date_str(rec.tanggalBl), get_str(house.nomorHostBl),
                    get_date_str(house.tanggalHostBl) or get_date_str(rec.tanggalBl), get_str(rec.namaSaranaPengangkut),
                    get_str(house.npwpPenerima) or "0000000000000000", get_str(house.namaPenerima), get_str(house.alamatPenerima),
                    get_ref_name(house.negaraPenerima) or "ID", get_str(house.id_shipper), get_str(house.namaPengirim),
                    get_str(house.alamatPengirim), get_ref_name(house.negaraPengirim),
                    get_str(house.namaNotify) or "SAME AS CONSIGNEE", get_str(house.alamatNotify) or "SAME AS CONSIGNEE",
                    "ID", get_ref_name(rec.kodePelabuhanAsal), get_ref_name(rec.kodePelabuhanTransit) or get_ref_name(rec.kodePelabuhanAsal),
                    get_ref_name(rec.kodePelabuhanBongkar), get_ref_name(rec.kodePelabuhanBongkar),
                    house.jumlahKemasan or 0, get_str(house.jenisKemasan) or "PK", get_str(house.marking) or "N/M",
                    tot_container, house.berat or 0.0, house.dimensi or 0.0, "N", house.jumlahKemasan or 0, tot_container, "",
                    # Menarik Jenis ID
                    get_ref_name(getattr(house, 'jenis_id_pengirim', False)),
                    get_ref_name(getattr(house, 'jenis_id_penerima', False)) or get_ref_name(getattr(house, 'jenis_id_consignee', False)),
                    
                    get_str(house.telp_penerima), 
                    get_str(house.telp_pengirim),
                    "", # NO BARCODE KOSONG
                    get_str(house.no_invoice), 
                    get_date_str(house.tgl_invoice), 
                    
                    get_str(getattr(house, 'jenis_aju', '')), 
                    get_str(getattr(house, 'jenis_pibk', '')), 
                    get_str(getattr(house, 'no_sub_sub_pos', '')) if getattr(house, 'no_sub_sub_pos', '') else "0000", 
                    get_str(getattr(house, 'kategori_barang', '')), 
                    "", ""
                ]
                for col_idx, val in enumerate(row_d):
                    ws_detil.write(idx+1, col_idx, val)

            # ==========================================
            # SHEET 4: MASTER ENTRY
            # ==========================================
            ws_master = workbook.add_sheet('Master Entry')
            cols_master = ['ID MASTER', 'NOMOR AJU', 'KD KELOMPOK POS', 'NO MASTER BL/AWB', 'TGL MASTER BL/AWB', 'JML HOST BL/AWB']
            for col_idx, col_name in enumerate(cols_master):
                ws_master.write(0, col_idx, col_name)
            
            row_m = [str(rec.id), get_str(rec.nomorAju), "01", get_str(rec.name), get_date_str(rec.tanggalBl), rec.jumlahPos]
            for col_idx, val in enumerate(row_m):
                ws_master.write(1, col_idx, val)

            # ==========================================
            # SHEET 5: HEADER
            # ==========================================
            ws_header = workbook.add_sheet('Header')
            cols_header = ['NPWP', 'JNS MANIFEST', 'KD JNS MANIFEST', 'KPPBC', 'NO BC 11', 'TGL BC 11', 'NAMA SARANA ANGKUT', 'KODE MODA', 'CALL SIGN', 'NO IMO', 'NEGARA', 'NO VOYAGE / ARRIVAL', 'TGL TIBA', 'JAM TIBA', 'TOTAL POS', 'TOTAL KEMASAN', 'TOTAL KONTAINER', 'TOTAL MASTER BL/AWB', 'TOTAL BRUTO', 'TOTAL VOLUME', 'PEMBERITAHU']
            for col_idx, col_name in enumerate(cols_header):
                ws_header.write(0, col_idx, col_name)
            
            row_h = [
                get_str(rec.idPerusahaan)[:15], 
                get_str(rec.jenisManifes) or "11", "EM", 
                get_ref_name(rec.kodeKantor),
                get_str(rec.nomor_bc11),       
                get_date_str(rec.tanggal_bc11), 
                get_str(rec.namaSaranaPengangkut), "1" if rec.modePengangkut == '1' else get_str(rec.modePengangkut),
                get_str(rec.callSign), get_str(rec.imoNumber), get_ref_name(rec.kodeNegara), get_str(rec.nomorVoyage),
                get_date_str(rec.tanggalTiba), get_time_str(rec.tanggalTiba), rec.jumlahPos, tot_kemasan,
                tot_container, 1, tot_bruto, tot_volume, get_str(rec.namaPerusahaan)
            ]
            for col_idx, val in enumerate(row_h):
                ws_header.write(1, col_idx, val)
            # ==========================================
            # SIMPAN KE BENTUK FILE BINARY
            # ==========================================
            fp = io.BytesIO()
            workbook.save(fp)
            fp.seek(0)
            data = fp.read()
            fp.close()
            
            file_name = "Form_B_%s.xls" % get_str(rec.nomorAju)
            rec.write({
                'form_b_file': base64.b64encode(data),
                'form_b_filename': file_name
            })

    # --- PROTEKSI CRUD ---
    @api.multi
    def unlink(self):
        for rec in self:
            if rec.state != 'draft':
                raise UserError("Data terkunci! Tidak bisa menghapus data yang statusnya bukan Draft.")
        return super(NvoccMaster, self).unlink()

    @api.multi
    def write(self, vals):
        for rec in self:
            allowed_fields = [
                'state', 'jumlahPos', 'json_file', 'json_filename', 'xml_file', 'xml_filename', 
                'nomorAju', 'form_b_file', 'form_b_filename', 
                'nomor_bc11', 'tanggal_bc11', 'nomor_pos_bc11'
            ]
            if rec.state != 'draft' and not any(k in vals for k in allowed_fields):
                raise UserError("Data terkunci! Klik 'Set to Draft' jika ingin mengubah data asli manifest.")
        return super(NvoccMaster, self).write(vals)

    def action_confirm(self):
        self.action_generate_json()

    def action_send_ceisa(self):
        for rec in self:
            rec.state = 'sent'

    def action_done(self):
        for rec in self:
            rec.state = 'done'

    def action_draft(self):
        for rec in self:
            rec.state = 'draft'

    def _get_default_kantor(self):
        kantor = self.env['nvocc.reference'].search([('name', '=', '060100'), ('kode_master', '=', 7)], limit=1)
        return kantor.id if kantor else False

    kodeKantor = fields.Many2one('nvocc.reference', string='Kode Kantor', 
                                domain=[('kode_master','=',7)], 
                                default=_get_default_kantor) 