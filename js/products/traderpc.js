/* ============================================================
   Trader PC v2 — option metadata
   --------------------------------------------------------------
   Keyed by `idoptoptgrp` (the unique row in options_optionsGroups
   for this product/option pair). Sparse fields per option — fill
   in only what the option actually needs. If a key is missing
   here, the page falls back to the DB description and skips
   ratings.

   Field reference (all optional unless noted):

     name             friendly display name (overrides the DB
                      description on the option button, the
                      cfg-row__selected caption, and the summary
                      list — wins everywhere it shows on screen)

     hide             true to remove this option from the page
                      entirely (button is dropped from the DOM
                      on load). If the hidden option happens to
                      be the group's default, the next remaining
                      option is promoted to selected. Omit or
                      set false to show normally.

     -- Full Specification panel --
     specKey          which row in the Full Spec table this option
                      drives. Required rows: 'cpu', 'ram', 'gpu',
                      'storage', 'os', 'warranty'. Optional rows
                      (only show when an option drives them):
                      'wifi', 'bluetooth', 'office', 'backup',
                      'extras'. When set, the option's specText
                      (or name) populates that row.
     specText         full text for the spec row (overrides `name`
                      for spec-row purposes only — option button
                      and summary list keep using `name`).
     specRows         { key: text, ... } for an option that drives
                      multiple spec rows at once. Use this when a
                      single option covers two categories — e.g. a
                      "Wifi 6 with Bluetooth" combo card sets both
                      the wifi row and the bluetooth row. Mutually
                      exclusive with specKey/specText on the same
                      entry; specRows wins.
     specSkip         true to hide the spec row(s) when this option
                      is selected — useful for an explicit "None"
                      choice. Equivalent to omitting specKey, but
                      documents intent. (Optional rows reset to
                      hidden each render automatically, so the
                      simpler pattern is to just leave specKey/
                      specRows unset on a "None" option.)

     ramShort,
     storageShort     short forms for the "Your build" summary line
                      under the spec table (e.g. "16 GB", "500 GB
                      NVMe"). Falls back to `name` if not set.

     -- CPU rows --
     coresLabel       display text for the inline CPU info card's
                      "Cores & threads" row, e.g. '10 cores / 16
                      threads'. Drives [data-cpu-stat="cores"].
     cpuSpeed         1-5 star rating (drives the "CPU Speed" rows in
                      both the sidebar [data-rating="speed"] and the
                      inline accordion card [data-cpu-stat="speed"])
     cpuMultiTask     1-5 star rating (bumped by RAM bonus before
                      driving [data-rating="mt"] and [data-cpu-stat="mt"])
     cpuMultiThread   1-5 star rating (drives the inline CPU info
                      card's Multi-Threaded row [data-cpu-stat="mthread"])

     cooler           full text for the CPU-cooler spec row, e.g.
                      'be quiet! Pure Rock 2 silent tower'
     coolerUpgraded   true to add the orange "auto-upgrade" badge
     mobo             full text for the motherboard spec row
     moboUpgraded     true to add the badge

     -- GPU rows --
     vram             e.g. '8 GB GDDR7'. Also drives the inline GPU
                      info card's "Video memory" row [data-gpu-stat="vram"].
     outputs          e.g. '4× DisplayPort, 1× HDMI'. Also drives the
                      inline GPU info card's "Display outputs" row
                      [data-gpu-stat="ports"].
     gpuPower         1-5 star rating (drives both the sidebar
                      [data-rating="gfx"] and the inline accordion
                      card [data-gpu-stat="gfx"])
     gpuAi            1-5 star rating (drives both the sidebar
                      [data-rating="ai"] and the inline accordion
                      card [data-gpu-stat="ai"])
     screens          number of screens this GPU supports (e.g.
                      4, 6, 8). When set, GPU options on the page
                      are sorted by this value (then by price) and
                      grouped under "{N} Screen Options" sub-
                      headings. Options without screens fall into
                      an "Other Options" bucket at the bottom.
     gpuLabel         short label for the right-side GPU panel
                      context line, e.g. 'RTX 5050 · 8 screens'
     resolutions      full text line of available monitor resolutions.
     psu              full text for the PSU spec row (driven by GPU
                      because high-power GPUs need a bigger PSU)
     psuUpgraded      true to add the auto-upgrade badge

     -- RAM rows --
     ramMtBonus       0 or 1, added to CPU multi-task rating

   ----------------------------------------------------------------
   Replace the PLACEHOLDER ids below with real `idoptoptgrp`
   values from the DB. To find them quickly: load the live page,
   inspect any option button — its `data-idoptoptgrp` attribute
   is the key you paste in here.
   ============================================================ */
window.MM_OPTION_META = {

  /* ---------- CPU ---------- */
  '18464': {
    name:           'Intel i5 14400F',
    coresLabel:     '10 Cores / 16 Threads / 2.5 - 4.7 GHz',
    cpuSpeed:       5.5,
    cpuMultiTask:   5,
    cpuMultiThread: 3.5,
    specKey:        'cpu',
    specText:       'Intel i5 14400F · 10C/16T',
    cooler:         'Be Quiet! Silent Cooler',
    coolerUpgraded: false,
    mobo:           'Fast B760 Chipset Motherboard',
    moboUpgraded:   false
  },
  '18479': {
    name:           'Intel i5 14600KF',
    coresLabel:     '14 Cores / 20 Threads / 3.5 - 5.3 GHz',
    cpuSpeed:       7,
    cpuMultiTask:   6.5,
    cpuMultiThread: 5.5,
    specKey:        'cpu',
    specText:       'Intel i5 14600KF · 14C/20T',
    cooler:         'Be Quiet! Silent Cooler - Enhanced',
    coolerUpgraded: true,
    mobo:           'Fast B760 Chipset Motherboard',
    moboUpgraded:   false
  },
  '18480': {
    name:           'Intel i7 14700KF',
    coresLabel:     '20 Cores / 28 Threads / 3.4 - 5.6 GHz',
    cpuSpeed:       7.5,
    cpuMultiTask:   9,
    cpuMultiThread: 7.5,
    specKey:        'cpu',
    specText:       'Intel i7 14700KF · 20C/28T',
    cooler:         'Be Quiet! Silent Cooler - Enhanced',
    coolerUpgraded: true,
    mobo:           'Fast B760 Chipset Motherboard',
    moboUpgraded:   false
  },
  '18496': {
    name:           'Intel i9 14900KF',
    coresLabel:     '24 Cores / 32 Threads / 3.2 - 6.0 GHz',
    cpuSpeed:       8.5,
    cpuMultiTask:   10,
    cpuMultiThread: 8.5,
    specKey:        'cpu',
    specText:       'Intel i9 14900KF · 24C/32T',
    cooler:         'Be Quiet! Silent Cooler - Enhanced',
    coolerUpgraded: true,
    mobo:           'Fast B760 Chipset Motherboard',
    moboUpgraded:   false
  },

  /* ---------- RAM ---------- */
  '17950': {
    name:       '16 GB DDR4 3200',
    ramMtBonus: 0,
    specKey:    'ram',
    specText:   '16 GB DDR4 3200MHz',
    ramShort:   '16 GB'
  },
  '17951': {
    name:       '32 GB DDR4 3200',
    ramMtBonus: 1,
    specKey:    'ram',
    specText:   '32 GB DDR4 3200MHz',
    ramShort:   '32 GB'
  },
  '18064': {
    name:       '64 GB DDR4 3200',
    ramMtBonus: 1,
    specKey:    'ram',
    specText:   '64 GB DDR4 3200MHz',
    ramShort:   '64 GB'
  },

  /* ---------- GPU ---------- */
  '18518': {
    name:         'nVidia RTX A400',
    screens:      4,
    vram:         '4GB',
    outputs:      '4x Mini-DisplayPort',
    gpuPower:     4,
    gpuAi:        43,
    gpuLabel:     'nVidia RTX A400 · 4 screens',
    resolutions:  '2x 4K, 4x QHD, 4x Full HD',
    specKey:      'gpu',
    specText:     'nVidia RTX A400 · 4GB · 4 screens',
    psu:          'Be Quiet! 550W',
    psuUpgraded:  false
  },
  '18510': {
    name:         'nVidia RTX 5050',
    screens:      4,
    vram:         '8GB',
    outputs:      '1x HDMI, 3x DisplayPort',
    gpuPower:     8.5,
    gpuAi:        421,
    gpuLabel:     'nVidia RTX 5050 · 4 screens',
    resolutions:  '4x 5K, 4x 4K, 4x QHD, 4x Full HD',
    specKey:      'gpu',
    specText:     'nVidia RTX 5050 · 8 GB · 4 screens',
    psu:          'Be Quiet! 550W',
    psuUpgraded:  false
  },
  '18519': {
    name:         'nVidia RTX A400 x2',
    screens:      8,
    vram:         '4GB Per Card',
    outputs:      '8x Mini-DisplayPort',
    gpuPower:     4,
    gpuAi:        43,
    gpuLabel:     'nVidia RTX A400 x2 · 8 screens',
    resolutions:  '2x 4K, 8x QHD, 8x Full HD',
    specKey:      'gpu',
    specText:     'nVidia RTX A400 x2 · 4GB · 8 screens',
    psu:          'Be Quiet! 550W',
    psuUpgraded:  false,
     mobo:         'Fast B760 Chipset Motherboard',
    moboUpgraded: true
  },
  '18511': {
    name:         'nVidia RTX 5050 x2',
    screens:      8,
    vram:         '8GB Per Card',
    outputs:      '2x HDMI, 6x DisplayPort',
    gpuPower:     8.5,
    gpuAi:        421,
    gpuLabel:     'nVidia RTX 5050 x2 · 8 screens',
    resolutions:  '4x 5K, 6x 4K, 8x QHD, 8x Full HD',
    specKey:      'gpu',
    specText:     'nVidia RTX 5050 x2 · 8 GB · 8 screens',
    psu:          'Be Quiet! 750W',
    psuUpgraded:  true,
    mobo:         'Fast B760 Chipset Motherboard',
    moboUpgraded: true
  },

  /* ---------- Boot Drive ---------- */
  '18393': {
    name:          '500GB NVMe SSD (Adata)',
    specKey:       'storage',
    specText:      '500 GB Adata NVMe M.2 · 3,500 MB/s read',
    storageShort:  '500 GB NVMe'
  },
  '18390': {
    name:          '1TB NVMe SSD (Kingston)',
    specKey:       'storage',
    specText:      '1TB Adata NVMe M.2 · 6,000 MB/s read',
    storageShort:  '1TB NVMe'
  },
  '18392': {
    name:          '2TB NVMe SSD (WD Blue)',
    specKey:       'storage',
    specText:      '2TB WD Blue NVMe M.2 · 6,000 MB/s read',
    storageShort:  '2TB NVMe'
  },
  '18465': {
    name:          '4TB NVMe SSD (Kingston)',
    specKey:       'storage',
    specText:      '4TB Kingston NVMe M.2 · 3,500 MB/s read',
    storageShort:  '4TB NVMe'
  },

  /* ---------- 2ND Drive ---------- */
  '18235': {
    name:          '1TB NVMe SSD (Adata)',
    specKey:       '2nddrive',
    specText:      '1TB Adata NVMe M.2 · 3,500 MB/s read',
    storageShort:  '1TB NVMe'
  },
  '18236': {
    name:          '2TB NVMe SSD (Adata)',
    specKey:       '2nddrive',
    specText:      '2TB Adata NVMe M.2 · 3,500 MB/s read',
    storageShort:  '2TB NVMe'
  },
  '18315': {
    name:          '4TB NVMe SSD (Kingston)',
    specKey:       '2nddrive',
    specText:      '4TB Kingston NVMe M.2 · 3,500 MB/s read',
    storageShort:  '4TB NVMe'
  },
  '18237': {
    name:          '4TB Traditional Style',
    specKey:       '2nddrive',
    specText:      '4TB Traditional Style HDD · 5,400RPM',
    storageShort:  '4TB HDD'
  },
  '18233': {
    name:          '6TB Traditional Style',
    specKey:       '2nddrive',
    specText:      '6TB Traditional Style HDD · 5,400RPM',
    storageShort:  '6TB HDD'
  },
  '18484': {
    name:          '8TB Traditional Style',
    specKey:       '2nddrive',
    specText:      '8TB Traditional Style HDD · 5,400RPM',
    storageShort:  '8TB HDD'
  },

  /* ---------- OS ---------- */
  '18443': {
    name:     'Windows 11 Home',
    specKey:  'os',
    specText: 'Windows 11 Home · pre-activated'
  },
  '18280': {
    name:     'Windows 11 Pro',
    specKey:  'os',
    specText: 'Windows 11 Pro · pre-activated'
  },

  /* ---------- Wifi / Bluetooth ----------
     The "Wifi 6 with Bluetooth" combo card drives BOTH the
     wifi spec row and the bluetooth spec row via specRows.
     The "None" choice has no specKey, so both rows stay
     hidden when it's selected. */
  '18133': { hide: true },
  '17966': {
    name:     'Wifi 6 with BlueTooth',
    specRows: {
      wifi:      'Wi-Fi 6 (802.11ax) · dual-band',
      bluetooth: 'Bluetooth 5.2'
    }
  },
  '18112': { name: 'None' },

  /* ---------- Microsoft Office ----------
     Replace placeholder ids with real `idoptoptgrp` values
     once Office is added as an option group on the product. */
  '17906': { name: 'None' },
  '18063': {
    name:     'Home 2024',
    specKey:  'office',
    specText: 'Microsoft Office Home 2024 · lifetime licence'
  },
  '18062': {
    name:     'Home & Business 2024',
    specKey:  'office',
    specText: 'Microsoft Office Home & Business 2024 · lifetime licence'
  },

  /* ---------- Backup System ---------- */
  '18040': { name: 'None' },
  '17915': {
    name:     'DVD ReWriter',
    specKey:  'optical',
    specText: 'DVD ReWriter'
  },

  /* ---------- Inputs ---------- */
  '18113': { name: 'None' },
  '18114': {
    name:     'Wired Keyboard / Mouse',
    specKey:  'inputs',
    specText: 'Logitech Wired Keyboard / Mouse Set'
  },
  '17894': {
    name:     'Wireless Keyboard / Mouse',
    specKey:  'inputs',
    specText: 'Logitech Wireless Keyboard / Mouse Set'
  },

  /* ---------- Speakers ---------- */
  '18132': { hide: true },
  '18111': { name: 'None' },
  '17897': {
    name:     'Desktop Speakers',
    specKey:  'speakers',
    specText: 'Logitech USB desktop speakers'
  },

  /* ---------- Optical ---------- */
  '18245': { name: 'None' },
  '18039': {
    name:     'Bootable Backup Drive',
    specKey:  'backup',
    specText: 'Bootable backup hard drive · Instant Windows recovery'
  },

  /* ---------- bluetooth ---------- */
  '18246': { name: 'None' },
  '18247': {
    name:     'USB Bluetooth Adapter',
    specKey:  'bluetooth',
    specText: 'USB Bluetooth Adapter'
  },

  /* ---------- Warranty ---------- */
  '17992': {
    name:     '1 Yr OnSite / 5 Yr Labour',
    specKey:  'warranty',
    specText: '5-year hardware cover · 1-year OnSite'
  },
  '17903': {
    name:     '2 Yr OnSite / 5 Yr Labour',
    specKey:  'warranty',
    specText: '5-year hardware cover · 2-year OnSite'
  },
  '17905': {
    name:     '3 Yr OnSite / 5 Yr Labour',
    specKey:  'warranty',
    specText: '5-year hardware cover · 3-year OnSite'
  }

};
