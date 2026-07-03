/* =============================================================
   TEREX INVENTORY — CUSTOMER SCHEMA DEFINITIONS
   Tiap customer mendefinisikan:
   - id, name, color
   - sheetInbound / sheetOutbound : nama sheet di Excel (case-insensitive)
   - headerRowDetect : bagaimana mendeteksi baris header
   - map : pemetaan kolom Excel → field standar internal
   - statusReady / statusFaulty : nilai status yang dianggap ready/faulty
   - isUnreturned(row) : apakah baris outbound belum dikembalikan ke WH
   - isTechnicianPIC(val) : apakah nilai PIC termasuk teknisi sungguhan
   - inboundTechLocField : kolom di Inbound yang berisi nama teknisi (untuk stock ready)
============================================================= */

const CUSTOMERS = {

  /* ------------------------------------------------------------------ */
  IPT: {
    id: 'IPT',
    name: 'IPT',
    color: '#E31E24',
    /* Google Sheets CSV URLs — isi setelah publish sheet ke web.
       Format: https://docs.google.com/spreadsheets/d/[ID]/gviz/tq?tqx=out:csv&sheet=[NAMA_SHEET]
       Kosongkan string ('') jika belum dikonfigurasi. */
    sheets: {
      inbound:  '',   // TODO: paste URL sheet Inbound IPT
      outbound: '',   // TODO: paste URL sheet Outbound IPT
    },
    sheetInbound:  'inbound',
    sheetOutbound: 'outbound',
    map: {
      /* field standar       : nama kolom di Excel (lowercase) */
      materialName:    'material name',
      itemCode:        'item code',
      qty:             'qty',
      serialNumber:    'serial number',
      status:          'status',
      pic:             'pic',            // di outbound
      lokasi:          'lokasi',         // di inbound (nama teknisi/site)
      noReturn:        'no ba',          // di outbound — kosong = belum return
      region:          'region',
      noBKB:           'no bkb',
      kebutuhan:       'kebutuhan',
      kepemilikan:     'kepemilikan',
      supplier:        'supplier',
      asalPickup:      'asal pickup',
      tanggalPickup:   'tanggal pickup',
      tanggalInstall:  'tanggal install',
      tanggalReturn:   'tanggal return',
      numberTT:        'number tt',
      remark:          'remark',
      updateDate:      'update date',
      noBA:            'no ba',
      asalSite:        'asal site',
      updateNote:      'update note',
    },
    statusReady(s){
      const v = String(s||'').toLowerCase().trim();
      return v.includes('ready') || v === 'baik';
    },
    statusFaulty(s){
      const v = String(s||'').toLowerCase().trim();
      return v.includes('fault') || v.includes('rusak') || v.includes('damage') || v.includes('reject') || v.includes('bad');
    },
    isUnreturned(row){ return !String(row['no ba']||'').trim(); },
    NON_TECH_PIC: new Set(['no need','on search','fariz']),
    isTechnicianPIC(val){
      const v = String(val||'').trim().toLowerCase();
      return !!v && !this.NON_TECH_PIC.has(v);
    },
    inboundTechLocField: 'lokasi',

    /* Kolom yang ditampilkan di tabel Inbound */
    inboundColumns: [
      { key:'material name',   label:'Material Name' },
      { key:'item code',       label:'Item Code',    mono:true },
      { key:'qty',             label:'Qty',          num:true  },
      { key:'serial number',   label:'Serial Number',mono:true },
      { key:'kebutuhan',       label:'Kebutuhan'               },
      { key:'no bkb',          label:'No BKB',       mono:true },
      { key:'kepemilikan',     label:'Kepemilikan'             },
      { key:'supplier',        label:'Supplier'                },
      { key:'asal pickup',     label:'Asal Pickup'             },
      { key:'tanggal pickup',  label:'Tgl Pickup',   date:true },
      { key:'tanggal install', label:'Tgl Install',  date:true },
      { key:'lokasi',          label:'Lokasi'                  },
      { key:'region',          label:'Region'                  },
      { key:'status',          label:'Status',       badge:true},
      { key:'remark',          label:'Remark'                  },
      { key:'update date',     label:'Update Date',  date:true },
    ],
    /* Kolom yang ditampilkan di tabel Outbound */
    outboundColumns: [
      { key:'material name',   label:'Material Name' },
      { key:'item code',       label:'Item Code',    mono:true },
      { key:'qty',             label:'Qty',          num:true  },
      { key:'serial number',   label:'Serial Number',mono:true },
      { key:'asal site',       label:'Asal Site'               },
      { key:'status',          label:'Status',       badge:true},
      { key:'no ba',           label:'No BA',        mono:true },
      { key:'tanggal return',  label:'Tgl Return',   date:true },
      { key:'pic',             label:'PIC'                     },
      { key:'remark',          label:'Remark'                  },
      { key:'update note',     label:'Update Note'             },
    ],
    inboundFilterRegionField: 'region',
    outboundReturnField: 'no ba',   // kosong = belum return
    outboundReturnType: 'empty',    // 'empty' | 'date'
    searchKeysInbound:  ['material name','item code','serial number','lokasi','region','no bkb'],
    searchKeysOutbound: ['material name','item code','serial number','asal site','pic'],
  },

  /* ------------------------------------------------------------------ */
  MSG: {
    id: 'MSG',
    name: 'MSG',
    color: '#1A6FC4',
    sheets: {
      inbound:  'https://docs.google.com/spreadsheets/d/e/2PACX-1vS9qfQ_Vyl6BUMhfsdgdo2ACvYimblokz70Ti8rW_UW2l_q8_8r_hpjEapt3sN8bQ/pub?gid=2072440662&single=true&output=csv',
      outbound: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vS9qfQ_Vyl6BUMhfsdgdo2ACvYimblokz70Ti8rW_UW2l_q8_8r_hpjEapt3sN8bQ/pub?gid=1646686255&single=true&output=csv',
    },
    sheetInbound:  '02.data barang masuk',
    sheetOutbound: '03. data barang keluar',
    map: {
      materialName:    'nama barang',
      itemCode:        'kode barang',
      qty:             'jumlah',
      serialNumber:    'no sn',
      status:          'status perangkat',   // di inbound
      statusOut:       'status',             // di outbound
      pic:             'pic eos',
      lokasi:          'site',
      tanggalPickup:   'tanggal pickup',
      tanggalInstall:  'tanggal install',
      tanggalReturn:   'tanggal return',
      tanggalFaulty:   'tanggal faulty',
      diTerima:        'di terima',
      diSerahkan:      'di serahkan',
      warehouseDiterima: 'warehouse diterima',
      exSite:          'ex site',
      kabupaten:       'kabupaten',
      keteranganPerangkat: 'keterangan perangkat',
      satuan:          'satuan',
      problem:         'problem',
    },
    statusReady(s){
      const v = String(s||'').toLowerCase().trim();
      return v === 'ready';
    },
    statusFaulty(s){
      const v = String(s||'').toLowerCase().trim();
      return v === 'faulty';
    },
    isUnreturned(row){
      /* Di outbound MSG: TANGGAL RETURN kosong = belum dikembalikan */
      return !row['tanggal return'];
    },
    NON_TECH_PIC: new Set(['pak abe','pak chandra','pak rifqy','pak haryo','tsa','fariz']),
    isTechnicianPIC(val){
      const v = String(val||'').trim().toLowerCase();
      return !!v && !this.NON_TECH_PIC.has(v);
    },
    inboundTechLocField: 'pic eos',  // di inbound MSG, kolom PIC EOS = nama teknisi

    inboundColumns: [
      { key:'nama barang',          label:'Nama Barang'                 },
      { key:'kode barang',          label:'Kode Barang',   mono:true    },
      { key:'jumlah',               label:'Qty',           num:true     },
      { key:'satuan',               label:'Satuan'                      },
      { key:'no sn',                label:'No SN',         mono:true    },
      { key:'di terima',            label:'Di Terima'                   },
      { key:'di serahkan',          label:'Di Serahkan'                 },
      { key:'warehouse diterima',   label:'Warehouse'                   },
      { key:'tanggal pickup',       label:'Tgl Pickup',    date:true    },
      { key:'tanggal install',      label:'Tgl Install',   date:true    },
      { key:'pic eos',              label:'PIC EOS'                     },
      { key:'site',                 label:'Site'                        },
      { key:'status perangkat',     label:'Status',        badge:true   },
      { key:'keterangan perangkat', label:'Keterangan'                  },
    ],
    outboundColumns: [
      { key:'nama barang',       label:'Nama Barang'                  },
      { key:'kode barang',       label:'Kode Barang',    mono:true    },
      { key:'jumlah',            label:'Qty',            num:true     },
      { key:'no sn',             label:'No SN',          mono:true    },
      { key:'pic eos',           label:'PIC EOS'                      },
      { key:'di terima',         label:'Di Terima'                    },
      { key:'di serahkan',       label:'Di Serahkan'                  },
      { key:'status',            label:'Status',         badge:true   },
      { key:'tanggal faulty',    label:'Tgl Faulty',     date:true    },
      { key:'tanggal return',    label:'Tgl Return',     date:true    },
      { key:'ex site',           label:'Ex Site'                      },
      { key:'kabupaten',         label:'Kabupaten'                    },
    ],
    inboundFilterRegionField: null,          // MSG tidak punya kolom Region
    outboundReturnField: 'tanggal return',
    outboundReturnType: 'date',              // kosong = belum return
    searchKeysInbound:  ['nama barang','kode barang','no sn','site','pic eos'],
    searchKeysOutbound: ['nama barang','kode barang','no sn','ex site','pic eos','kabupaten'],
  },

  /* ------------------------------------------------------------------ */
  RGR: {
    id: 'RGR',
    name: 'RGR',
    color: '#1E8E5A',
    sheets: {
      inbound:  '',   // TODO: paste URL sheet Inbound RGR
      outbound: '',   // TODO: paste URL sheet Outbound RGR
    },
    sheetInbound:  '02.data barang masuk',
    sheetOutbound: '03. data barang keluar',
    map: {
      materialName:    'nama barang',
      itemCode:        'kode barang',
      qty:             'jumlah',
      serialNumber:    'no sn',
      status:          'status perangkat',   // di inbound: 'Ready' / 'Install'
      statusOut:       'status',             // di outbound: 'Faulty' / 'Return'
      pic:             'pic eos',
      lokasi:          'site',
      tanggalPickup:   'tanggal pickup',
      tanggalInstall:  'tanggal install',
      tanggalReturn:   'tanggal return',
      tanggalReplacement: 'tanggal replacement',
      tanggalPengiriman:  'tanggal pengiriman',
      diTerima:        'di terima',
      diSerahkan:      'di serahkan',
      warehouseDiterima: 'warehouse diterima',
      exSite:          'ex site',
      satuan:          'satuan',
      keterangan:      'keterangan perangkat',
      problem:         'problem',
    },
    statusReady(s){
      const v = String(s||'').toLowerCase().trim();
      return v === 'ready';
    },
    statusFaulty(s){
      const v = String(s||'').toLowerCase().trim();
      /* di outbound: 'faulty' = belum return; di inbound tidak ada faulty */
      return v === 'faulty';
    },
    isUnreturned(row){
      /* Gunakan kolom STATUS outbound: 'Faulty' = masih dipegang, belum return.
         'Return' = sudah dikembalikan (meski TANGGAL RETURN kadang kosong) */
      const s = String(row['status']||'').toLowerCase().trim();
      return s === 'faulty';
    },
    NON_TECH_PIC: new Set(['pic lokasi','alvin','pak firizqi','fariz']),
    isTechnicianPIC(val){
      const v = String(val||'').trim().toLowerCase();
      return !!v && !this.NON_TECH_PIC.has(v);
    },
    inboundTechLocField: 'pic eos',

    inboundColumns: [
      { key:'tanggal pickup',       label:'Tgl Pickup',    date:true    },
      { key:'nama barang',          label:'Nama Barang'                 },
      { key:'kode barang',          label:'Kode Barang',   mono:true    },
      { key:'jumlah',               label:'Qty',           num:true     },
      { key:'satuan',               label:'Satuan'                      },
      { key:'no sn',                label:'No SN',         mono:true    },
      { key:'di terima',            label:'Di Terima'                   },
      { key:'di serahkan',          label:'Di Serahkan'                 },
      { key:'warehouse diterima',   label:'Warehouse'                   },
      { key:'tanggal install',      label:'Tgl Install',   date:true    },
      { key:'pic eos',              label:'PIC EOS'                     },
      { key:'site',                 label:'Site'                        },
      { key:'status perangkat',     label:'Status',        badge:true   },
      { key:'keterangan perangkat', label:'Keterangan'                  },
    ],
    outboundColumns: [
      { key:'tanggal replacement',  label:'Tgl Replacement', date:true  },
      { key:'nama barang',          label:'Nama Barang'                 },
      { key:'kode barang',          label:'Kode Barang',   mono:true    },
      { key:'jumlah',               label:'Qty',           num:true     },
      { key:'no sn',                label:'No SN',         mono:true    },
      { key:'pic eos',              label:'PIC EOS'                     },
      { key:'di terima',            label:'Di Terima'                   },
      { key:'di serahkan',          label:'Di Serahkan'                 },
      { key:'status',               label:'Status',        badge:true   },
      { key:'tanggal pengiriman',   label:'Tgl Pengiriman', date:true   },
      { key:'tanggal return',       label:'Tgl Return',    date:true    },
      { key:'ex site',              label:'Ex Site'                     },
      { key:'problem',              label:'Problem'                     },
    ],
    inboundFilterRegionField: null,
    outboundReturnField: 'status',
    outboundReturnType: 'value',
    searchKeysInbound:  ['nama barang','kode barang','no sn','site','pic eos'],
    searchKeysOutbound: ['nama barang','kode barang','no sn','ex site','pic eos'],
  },

};

/* Daftar customer yang aktif (urutan tampil di dropdown) */
const CUSTOMER_LIST = ['IPT', 'MSG', 'RGR'];
/* Tambahkan 'PIM' di sini setelah format Excel-nya diketahui */
