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
     cores            e.g. '6P+4E'
     threads          e.g. 16
     ghz              e.g. '4.7'
     cpuSpeed         1-5 star rating (drives [data-rating="speed"])
     cpuMultiTask     1-5 star rating (further bumped by RAM bonus
                      to drive [data-rating="mt"])
     cpuMultiThread   1-5 star rating (reserved for inline CPU
                      info card — wired in step 2 of the rollout)

     cooler           full text for the CPU-cooler spec row, e.g.
                      'be quiet! Pure Rock 2 silent tower'
     coolerUpgraded   true to add the orange "auto-upgrade" badge
     mobo             full text for the motherboard spec row
     moboUpgraded     true to add the badge

     -- GPU rows --
     vram             e.g. '8 GB GDDR7'
     outputs          e.g. '4× DisplayPort, 1× HDMI'
     gpuPower         1-5 star rating (drives [data-rating="gfx"])
     gpuAi            1-5 star rating (drives [data-rating="ai"])
     screens          number of screens this GPU supports (e.g.
                      4, 6, 8). When set, GPU options on the page
                      are sorted by this value (then by price) and
                      grouped under "{N} Screen Options" sub-
                      headings. Options without screens fall into
                      an "Other Options" bucket at the bottom.
     gpuLabel         short label for the right-side GPU panel
                      context line, e.g. 'RTX 5050 · 8 screens'
     monitors         array of { count, res } feeding [data-mons]
                        e.g. [{ count: 8, res: '4K @ 120 Hz' }, ...]
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
    cores:          '6P+4E',
    threads:        16,
    ghz:            '4.7',
    cpuSpeed:       3,
    cpuMultiTask:   3,
    cpuMultiThread: 3,
    specKey:        'cpu',
    specText:       'Intel i5 14400F · 10C/16T',
    cooler:         'be quiet! Pure Rock 2 silent tower',
    coolerUpgraded: false,
    mobo:           'MSI PRO B760M-P DDR4',
    moboUpgraded:   false
  },
  '18479': {
    name:           'Intel i5 14600KF',
    cores:          '6P+8E',
    threads:        20,
    ghz:            '5.3',
    cpuSpeed:       4,
    cpuMultiTask:   3,
    cpuMultiThread: 4,
    specKey:        'cpu',
    specText:       'Intel i5 14600KF · 14C/20T',
    cooler:         'be quiet! Dark Rock 4 (120 mm tower)',
    coolerUpgraded: true,
    mobo:           'MSI PRO Z790-P DDR4',
    moboUpgraded:   true
  },
  'CPU_PLACEHOLDER_3': {
    name:           'Intel i7 14700KF',
    cores:          '8P+12E',
    threads:        28,
    ghz:            '5.6',
    cpuSpeed:       5,
    cpuMultiTask:   4,
    cpuMultiThread: 5,
    specKey:        'cpu',
    specText:       'Intel i7 14700KF · 20C/28T',
    cooler:         'be quiet! Dark Rock 4 (120 mm tower)',
    coolerUpgraded: true,
    mobo:           'MSI PRO Z790-P DDR4',
    moboUpgraded:   true
  },
  'CPU_PLACEHOLDER_4': {
    name:           'Intel i9 14900KF',
    cores:          '8P+16E',
    threads:        32,
    ghz:            '6.0',
    cpuSpeed:       5,
    cpuMultiTask:   5,
    cpuMultiThread: 5,
    specKey:        'cpu',
    specText:       'Intel i9 14900KF · 24C/32T',
    cooler:         'be quiet! Dark Rock Pro 5 (135 mm tower)',
    coolerUpgraded: true,
    mobo:           'MSI PRO Z790-P DDR4',
    moboUpgraded:   true
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
  '18466': {
    name:         'Intel Arc A380',
    screens:      4,
    vram:         '6 GB GDDR6',
    outputs:      '1x HDMI, 3x DisplayPort',
    gpuPower:     3,
    gpuAi:        2,
    gpuLabel:     'Intel A380 · 4 screens',
    monitors: [
      { count: 4, res: '4K @ 60 Hz'    },
      { count: 4, res: '1440p @ 144 Hz'},
      { count: 4, res: '1080p @ 240 Hz'}
    ],
    specKey:      'gpu',
    specText:     'Intel Arc A380 · 6 GB GDDR6 · 4 screens',
    psu:          'be quiet! Pure Power 12 500 W · 80+ Gold',
    psuUpgraded:  false
  },
  '18467': {
    name:         'Intel Arc A380 & Intel UHD',
    screens:      6,
    vram:         '6GB',
    outputs:      '2x HDMI, 4x DisplayPort',
    gpuPower:     3,
    gpuAi:        2,
    gpuLabel:     'Intel A380 & UHS · 6 screens',
    monitors: [
      { count: 6, res: '4K @ 60 Hz'    },
      { count: 6, res: '1440p @ 144 Hz'},
      { count: 6, res: '1080p @ 240 Hz'}
    ],
    specKey:      'gpu',
    specText:     'Intel Arc A380 + UHD · 6 GB · 6 screens',
    psu:          'be quiet! Pure Power 12 550 W · 80+ Gold',
    psuUpgraded:  true
  },
  'GPU_PLACEHOLDER_3': {
    name:         'NVIDIA RTX 5050',
    vram:         '8 GB GDDR7',
    outputs:      '4× DisplayPort, 1× HDMI',
    gpuPower:     5,
    gpuAi:        5,
    gpuLabel:     'RTX 5050 · 8 screens',
    monitors: [
      { count: 8, res: '4K @ 120 Hz'   },
      { count: 8, res: '1440p @ 240 Hz'},
      { count: 8, res: '1080p @ 360 Hz'}
    ],
    specKey:      'gpu',
    specText:     'NVIDIA RTX 5050 · 8 GB GDDR7 · 8 screens',
    psu:          'be quiet! Pure Power 12 650 W · 80+ Gold',
    psuUpgraded:  true
  },

  /* ---------- Storage ---------- */
  '18393': {
    name:          '500 GB Adata NVMe',
    specKey:       'storage',
    specText:      '500 GB Adata NVMe · M.2 · 3,500 MB/s read',
    storageShort:  '500 GB NVMe'
  },
  '18390': {
    name:          '1TB NVMe (Adata)',
    specKey:       'storage',
    specText:      '1 TB Adata NVMe · M.2 · 3,500 MB/s read',
    storageShort:  '1 TB NVMe'
  },
  'STORAGE_PLACEHOLDER_3': {
    name:          '2 TB NVMe Gen 4',
    specKey:       'storage',
    specText:      '2 TB NVMe Gen 4 · M.2 · 5,000 MB/s read',
    storageShort:  '2 TB NVMe'
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
    name:     'Home & Student 2024',
    specKey:  'office',
    specText: 'Microsoft Office Home & Student 2024 · lifetime licence'
  },
  '18062': {
    name:     'Home & Business 2024',
    specKey:  'office',
    specText: 'Microsoft Office Home & Business 2024 · lifetime licence'
  },

  /* ---------- Backup System ---------- */
  'BACKUP_PLACEHOLDER_NONE': { name: 'None' },
  'BACKUP_PLACEHOLDER_HDD': {
    name:     '2 TB backup HDD',
    specKey:  'backup',
    specText: '2 TB internal backup HDD · scheduled image backups'
  },
  'BACKUP_PLACEHOLDER_NAS': {
    name:     '4 TB external NAS backup',
    specKey:  'backup',
    specText: '4 TB external NAS · automated nightly backups'
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
  },

  /* ---------- Extras ---------- */
  'EXTRAS_PLACEHOLDER_1': {
    name:        'Cable management kit',
    specKey:     'extras',
    specText:    'Cable management kit'
  },
  'EXTRAS_PLACEHOLDER_2': {
    name:        'Premium thermal compound',
    specKey:     'extras',
    specText:    'Premium thermal compound'
  }

};
