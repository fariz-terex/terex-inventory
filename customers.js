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
      inbound:  'https://docs.google.com/spreadsheets/d/e/2PACX-1vRKpjd1obTGuY2dqhOFUXtrfh7ACVntlvSjz1firGklXv0KLqMHx13aUuaAvlU02Jhpx2_M3DYmycFC/pub?gid=774737978&single=true&output=csv',
      outbound: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRKpjd1obTGuY2dqhOFUXtrfh7ACVntlvSjz1firGklXv0KLqMHx13aUuaAvlU02Jhpx2_M3DYmycFC/pub?gid=1044779915&single=true&output=csv',
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
      inbound:  'https://docs.google.com/spreadsheets/d/e/2PACX-1vSBJjcfM6sfOpuIxvPo-txCoyxGlZV410Hd1XkiNjm0ccFVmv0bcSD1-7YBvs-ZWA/pub?gid=269731844&single=true&output=csv',
      outbound: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vSBJjcfM6sfOpuIxvPo-txCoyxGlZV410Hd1XkiNjm0ccFVmv0bcSD1-7YBvs-ZWA/pub?gid=615071854&single=true&output=csv',
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

  /* ------------------------------------------------------------------ */
  /* PIM — format BEDA dari customer lain: cuma 1 sheet gabungan
     ("Detail Terex"), tidak ada sheet Inbound/Outbound terpisah, dan
     tidak ada kolom PIC/teknisi (per-unit tracking berbasis Cluster/Site).
     Inbound  = baris dengan Status Material = AVAILABLE (stock di gudang)
     Outbound = baris dengan Status Material = Replacement (unit terpasang)
  ============================================================ */
  PIM: {
    id: 'PIM',
    name: 'PIM',
    color: '#8E44AD',
    /* Isi dengan link publish-to-web CSV dari sheet "Detail Terex".
       Karena inbound & outbound berasal dari sheet YANG SAMA, isi
       kedua URL ini dengan link yang SAMA persis. */
    sheets: {
      inbound:  '',
      outbound: '',
    },
    sheetInbound:  'Detail Terex',
    sheetOutbound: 'Detail Terex',
    /* single-sheet: setelah di-parse, baris dipisah otomatis berdasarkan
       Status Material menggunakan filterInboundRows / filterOutboundRows */
    singleSheet: true,
    map: {
      noRow:            'no',
      updateDate:       'date update',
      vendor:            'vendor',
      siteId:            'site id',
      lokasi:            'site name',     // dipakai juga sbg "lokasi" umum
      mrNumber:          'mr number',
      cluster:           'cluster',
      region:            'cluster',        // dipakai utk filter region di Inbound
      qty:               'qty shipment',
      itemCode:          'item code',
      materialName:      'material',
      brand:             'brand',
      type:              'type',
      serialNumber:      's/n after',
      serialNumberBefore:'s/n before',
      clusterActual:     'cluster actual',
      numberTT:          'tt number',
      noBKB:             'bak number',
      basNumber:         'bas number',
      qtyReplace:        'qty replace',
      justifikasi:       'justifikasi',
      typeBefore:        'type before',
      dateBAK:           'date bak',
      bakStatus:         'bak status',
      basStatus:         'bas status',
      tanggalInstall:    'date instalasi',
      tanggalReturn:     'date return',
      noBA:              'ba return number',
      noReturn:          'ba return number',   // konsisten dgn pola customer lain (kosong = belum return)
      status:            'status material',    // dipakai di Inbound (selalu AVAILABLE stlh filter)
      statusOut:         'status material',    // dipakai di Outbound (selalu Replacement stlh filter)
      statusReturnFaulty:'status returnfaulty',
      sohGudang:         'soh',
      sohTotal:          'soh_2',              // kolom SOH kedua (nama kolom sama di Excel, di-dedup jadi 'soh_2')
      issueSN:           'issue sn existing',
      pic:               'cluster',            // tidak ada PIC asli; dipakai hanya utk tampilan kartu Faulty Material di dashboard
    },
    statusReady(s){
      return String(s||'').toLowerCase().trim() === 'available';
    },
    statusFaulty(s){
      return String(s||'').toLowerCase().trim() === 'return spare faulty';
    },
    isUnreturned(row){
      /* Outbound (Replacement): belum return kalau BA Return Number kosong */
      return !String(row['ba return number']||'').trim();
    },
    NON_TECH_PIC: new Set(),
    isTechnicianPIC(){ return false; },  // PIM tidak punya konsep PIC/teknisi
    inboundTechLocField: null,

    /* ---- Khusus single-sheet: pisahkan hasil parse jadi Inbound/Outbound ---- */
    filterInboundRows(rows){
      return rows.filter(r => String(r['status material']||'').trim().toLowerCase() === 'available');
    },
    filterOutboundRows(rows){
      return rows.filter(r => String(r['status material']||'').trim().toLowerCase() === 'replacement');
    },

    /* ---- Halaman Stock: per Cluster (bukan per teknisi) ---- */
    stockMode: 'site',
    stockGroupField: 'cluster',
    stockGroupLabel: 'Cluster',
    stockPageTitle: 'Stock per Cluster',
    stockSearchPlaceholder: 'Cari nama cluster...',
    kpiTeknisiLabel: 'Total Cluster',
    computeKpiTeknisi(outboundRows){
      const set = new Set(outboundRows.map(r => String(r['cluster']||'').trim()).filter(Boolean));
      return set.size;
    },

    inboundColumns: [
      { key:'material',       label:'Material'                },
      { key:'item code',      label:'Item Code',   mono:true   },
      { key:'brand',          label:'Brand'                    },
      { key:'type',           label:'Type'                     },
      { key:'qty shipment',   label:'Qty',         num:true    },
      { key:'s/n after',      label:'Serial Number',mono:true  },
      { key:'cluster',        label:'Cluster'                  },
      { key:'vendor',         label:'Vendor'                   },
      { key:'mr number',      label:'MR Number',   mono:true   },
      { key:'status material',label:'Status',      badge:true  },
      { key:'date update',    label:'Update Date', date:true   },
    ],
    outboundColumns: [
      { key:'material',        label:'Material'                },
      { key:'item code',       label:'Item Code',   mono:true   },
      { key:'qty shipment',    label:'Qty',         num:true    },
      { key:'s/n after',       label:'S/N After',   mono:true   },
      { key:'s/n before',      label:'S/N Before',  mono:true   },
      { key:'site name',       label:'Site Name'                },
      { key:'cluster',         label:'Cluster'                  },
      { key:'tt number',       label:'TT Number',   mono:true   },
      { key:'date instalasi',  label:'Tgl Instalasi',date:true  },
      { key:'bak status',      label:'BAK Status'               },
      { key:'bas status',      label:'BAS Status'               },
      { key:'ba return number',label:'No BA Return',mono:true   },
      { key:'date return',     label:'Tgl Return',  date:true   },
      { key:'justifikasi',     label:'Justifikasi'              },
    ],
    /* Detail unit per cluster di halaman Stock */
    stockDetailColumns: [
      { key:'material',        label:'Material'               },
      { key:'item code',       label:'Item Code',  mono:true   },
      { key:'s/n after',       label:'S/N After',  mono:true   },
      { key:'site name',       label:'Site Name'               },
      { key:'tt number',       label:'TT Number',  mono:true   },
      { key:'bak status',      label:'BAK Status'              },
      { key:'bas status',      label:'BAS Status'              },
      { key:'date instalasi',  label:'Tgl Instalasi', date:true},
      { key:'ba return number',label:'No BA Return',mono:true  },
    ],
    inboundFilterRegionField: 'cluster',
    outboundReturnField: 'ba return number',
    outboundReturnType: 'empty',
    searchKeysInbound:  ['material','item code','s/n after','cluster','vendor','mr number'],
    searchKeysOutbound: ['material','item code','s/n after','s/n before','site name','cluster','tt number'],
  },

};

/* Daftar customer yang aktif (urutan tampil di dropdown) */
const CUSTOMER_LIST = ['IPT', 'MSG', 'RGR', 'PIM'];
