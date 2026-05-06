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

     -- CPU rows --
     cores            e.g. '6P+4E'
     threads          e.g. 16
     ghz              e.g. '4.7'
     cpuSpeed         1-5 star rating (drives [data-rating="speed"])
     cpuMultiTask     1-5 star rating (further bumped by RAM bonus
                      to drive [data-rating="mt"])
     cpuMultiThread   1-5 star rating (reserved for inline CPU
                      info card — wired in step 2 of the rollout)

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

     -- RAM rows --
     ramMtBonus       0 or 1, added to CPU multi-task rating

     -- Storage / OS / Warranty / Extras --
     name only — no other fields needed.

   ----------------------------------------------------------------
   Replace the PLACEHOLDER ids below with real `idoptoptgrp`
   values from the DB. To find them quickly: load the live page,
   inspect any option button — its `data-idoptoptgrp` attribute
   is the key you paste in here.
   ============================================================ */
window.MM_OPTION_META = {

  /* ---------- CPU ---------- */
  '18464': {
    name:           'Intel i5 14400FF',
    cores:          '6P+4E',
    threads:        16,
    ghz:            '4.7',
    cpuSpeed:       3,
    cpuMultiTask:   3,
    cpuMultiThread: 3
  },
  '18479': {
    name:           'Intel i5 14600KFF',
    cores:          '6P+8E',
    threads:        20,
    ghz:            '5.3',
    cpuSpeed:       4,
    cpuMultiTask:   3,
    cpuMultiThread: 4
  },
  'CPU_PLACEHOLDER_3': {
    name:           'Intel i7 14700KF',
    cores:          '8P+12E',
    threads:        28,
    ghz:            '5.6',
    cpuSpeed:       5,
    cpuMultiTask:   4,
    cpuMultiThread: 5
  },
  'CPU_PLACEHOLDER_4': {
    name:           'Intel i9 14900KF',
    cores:          '8P+16E',
    threads:        32,
    ghz:            '6.0',
    cpuSpeed:       5,
    cpuMultiTask:   5,
    cpuMultiThread: 5
  },

  /* ---------- RAM ---------- */
  'RAM_PLACEHOLDER_1': { name: '16 GB DDR5-5600', ramMtBonus: 0 },
  'RAM_PLACEHOLDER_2': { name: '32 GB DDR5-5600', ramMtBonus: 1 },
  'RAM_PLACEHOLDER_3': { name: '64 GB DDR5-5600', ramMtBonus: 1 },

  /* ---------- GPU ---------- */
  '18466': {
    name:     'Intel Arc A380',
    screens:  4,
    vram:     '6 GB GDDR6',
    outputs:  '1x HDMI, 3x DisplayPort',
    gpuPower: 3,
    gpuAi:    2,
    gpuLabel: 'Intel A380 · 4 screens',
    monitors: [
      { count: 4, res: '4K @ 60 Hz'    },
      { count: 4, res: '1440p @ 144 Hz'},
      { count: 4, res: '1080p @ 240 Hz'}
    ]
  },
  '18467': {
    name:     'Intel Arc A380 & Intel UHD',
    screens:  6,
    vram:     '6GB',
    outputs:  '2x HDMI, 4x DisplayPort',
    gpuPower: 3,
    gpuAi:    2,
    gpuLabel: 'Intel A380 & UHS · 6 screens',
    monitors: [
      { count: 6, res: '4K @ 60 Hz'    },
      { count: 6, res: '1440p @ 144 Hz'},
      { count: 6, res: '1080p @ 240 Hz'}
    ]
  },
  'GPU_PLACEHOLDER_3': {
    name:     'NVIDIA RTX 5050',
    vram:     '8 GB GDDR7',
    outputs:  '4× DisplayPort, 1× HDMI',
    gpuPower: 5,
    gpuAi:    5,
    gpuLabel: 'RTX 5050 · 8 screens',
    monitors: [
      { count: 8, res: '4K @ 120 Hz'   },
      { count: 8, res: '1440p @ 240 Hz'},
      { count: 8, res: '1080p @ 360 Hz'}
    ]
  },

  /* ---------- Storage ---------- */
  '18393': { name: '500 GB Adata NVMe' },
  '18390': { name: '1TB NVMe (Adata)'   },
  'STORAGE_PLACEHOLDER_3': { name: '2 TB NVMe Gen 4'   },

  /* ---------- OS ---------- */
  'OS_PLACEHOLDER_1': { name: 'Windows 11 Home' },
  'OS_PLACEHOLDER_2': { name: 'Windows 11 Pro'  },

  /* ---------- Wifi ---------- */
  '18133': { hide: true },
  'OS_PLACEHOLDER_2': { name: 'Windows 11 Pro'  },

  /* ---------- Warranty ---------- */
  'WARRANTY_PLACEHOLDER_1': { name: '5 years parts & labour'    },
  'WARRANTY_PLACEHOLDER_2': { name: '5 years + on-site year 1'  },

  /* ---------- Extras ---------- */
  'EXTRAS_PLACEHOLDER_1': { name: 'Cable management kit'       },
  'EXTRAS_PLACEHOLDER_2': { name: 'Premium thermal compound'   }

};
