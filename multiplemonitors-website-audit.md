# Multiple Monitors Ltd — Website Audit & Improvement Recommendations

**Audit date:** April 2026
**Last revised:** April 2026 — added two-door homepage architecture (4.1), revised `/trading-computers/` framing as primary trader destination (4.2), added detailed `/computers/` non-trader destination strategy (4.3), added `/bundles/` page strategy (4.5), added `/display-systems/` and final array page strategy (4.6), added `/stands/` Synergy Stand story strategy (4.7), refined `/pages/synergy-stand/` migration plan with specific content items (4.7.10), added product pages consistency strategy (4.8), added path-to-world-class section for `/trading-computers/` (9)

**Pages reviewed:**
- Homepage: https://www.multiplemonitors.co.uk/
- Trading Computers landing: https://www.multiplemonitors.co.uk/trading-computers/
- Computers category: https://www.multiplemonitors.co.uk/computers/
- Long-form trading computers page: https://www.multiplemonitors.co.uk/pages/trading-computers/
- Trader PC product page (sampled): https://www.multiplemonitors.co.uk/products/trader-pc/

**Goal:** Increase enquiries, conversion rate, and sales — without replatforming from ProductCart/classic ASP.

**Strategic context for homepage decisions:** Roughly 80% of sales are to traders, but the remaining 20% (CAD users, financial controllers, security operations, accountants, general multi-screen professionals) is real revenue that must not be alienated. The homepage must serve both audiences confidently, then route them to deeply specialised pages. Trader-specific specialisation lives on `/trading-computers/`, which becomes a comprehensive "trader destination" page rather than a brief landing page. The dedicated `tradingcomputers.co.uk` domain remains pointed at multiplemonitors.co.uk for now; revisit in 6-12 months once `/trading-computers/` has been brought to a world-class state and conversion data is available to inform the decision.

---

## 1. Overall Impression

The business has genuine, defensible advantages — 17+ years of specialisation, proprietary testing data via TraderSpec.com, BBC TV credit, strong Trustpilot reputation, and real technical authority shown in the writing. The problem is that almost none of that comes across with appropriate weight on the pages a first-time visitor lands on.

The site reads like the work of someone who has been incrementally adding to the same template since around 2015. A first-time visitor evaluating a £1,500–£3,000 purchase forms an impression of "small UK shop with deep technical knowledge but maybe a bit of a one-man-band." That impression doesn't match the buying decision being asked of them, and it leaves money on the table — particularly against US competitors and generic UK PC builders like Scan 3XS.

Most of the issues are about **presentation, prominence, and trust signals** rather than missing content. The underlying material is mostly there.

---

## 2. Strengths (Things to Protect and Amplify)

### 2.1 TraderSpec.com testing data is your single biggest differentiator
None of the UK competitors can credibly claim to have done independent benchmarking of trading software workloads. The single-thread vs multi-thread CPU charts on `/trading-computers/`, with explicit callouts of competitor CPU choices being inferior, is exactly the kind of authority-building content that wins customers. This advantage should be louder, not quieter.

### 2.2 The long-form `/pages/trading-computers/` page is genuinely good copywriting
The "I feel your pain" opener, the breakdown of why most retailers don't really specialise, the "build it yourself for £25/hour" calculation, and the "in an ideal world..." section are all classic direct-response techniques executed competently. This page would convert well — if it were updated, scannable, and modernised.

### 2.3 Bundle proposition is strong on paper
Free cables, free wifi/speakers, free delivery, automatic discount, plus everything-compatible-and-tested. Genuinely better than a customer piecing together a setup from Scan + Amazon + a stand vendor. (See section 5 — it's currently undersold.)

### 2.4 Star-rating product configurator tied to benchmark data
Telling customers what each spec change actually means for their real-world trading workload is something competitors can't easily copy. Most competitors just list specs; you translate them into outcomes.

### 2.5 Confident, customer-first technical voice
Talking traders out of unnecessary i9 upgrades or expensive RAM is the kind of thing that builds trust, earns referrals, and differentiates from upsell-heavy competitors.

### 2.6 Genuinely strong offer (lifetime support, 5-year cover, 30-day money-back)
US competitors typically offer 3 years. You offer more. This is barely communicated.

---

## 3. Critical Issues — Probably Costing Sales Right Now

These are the things to fix first. Roughly ordered by impact-to-effort ratio.

### 3.1 "©Copyright 2022" in the footer of every page
Serious credibility problem. A first-time visitor evaluating a £2,000 purchase will spot this and unconsciously discount your trustworthiness. It implies the business may be abandoned. Set this to dynamically render the current year — should be a 30-second classic ASP fix.

### 3.2 Trustpilot is invisible
You have a 5-star Trustpilot reputation but there is no Trustpilot badge, review widget, star rating, or even a link on any of the audited pages. This is your strongest piece of social proof and you're hiding it. Trustpilot Business plans include embeddable widgets — these should appear in hero areas, near "Add to basket" buttons, and in the footer. Even just the badge with star count would help.

### 3.3 Pricing inconsistencies across pages
- `/pages/trading-computers/` long-form page lists Trader PC at £895
- `/trading-computers/` newer page says "From £1,045"
- `/products/trader-pc/` product page says £995 + VAT (£1,194 inc VAT)

A customer cross-referencing will spot this and either email to ask, lose confidence, or bounce. The long-form page in particular has copy that's clearly years old in places ("our brand new high performance trading computer" describing the Trader Pro, when it's clearly been around a while).

**Action:** Either have all pages pull pricing dynamically from a single source, or do a one-time reconciliation pass and add a note on edit dates.

### 3.4 BBC mention is buried in a tiny banner
"As seen on BBC's Traders: Millions by the Minute" is mentioned in a small banner most visitors will never read. This is a massive trust signal that should be much more prominent — ideally with a small still image or BBC logo treatment in the hero area of landing pages.

### 3.5 Testimonials are weak and identical across pages
The same three customer quotes (Tom Boszko, Geoff Wheeler, David Coomber) appear on multiple pages. They're short, surnames-only, with no photos, no dates, no business context, and don't reflect the kind of detailed Trustpilot reviews you presumably have. Three sentences from three customers from years ago is not credible social proof for a 17-year-old business that has supplied hedge funds.

### 3.6 No real photographs anywhere
Every image is a rendered PC product shot, a stock illustration, or a graphic. No photos of:
- Your workshop
- Your testing rig
- You (Darren) — putting a face to the business is huge for trust
- Customers' actual setups
- The build process
- The delivery boxes / unboxing

For a custom-build business, this is a missed opportunity to communicate craftsmanship and authenticity. A 30-second iPhone video of a stress test running on a build, or a wide shot of the workshop, would do enormous work.

### 3.7 No video content
Display Fusion video is mentioned on the long-form page but I don't see it embedded. Video on product pages — even just a 60-second walkthrough of the Trader PC, you talking about why it's spec'd the way it is — would be hugely valuable for conversion.

---

## 4. Page-by-Page Notes

### 4.1 Homepage — "Two-Door" Architecture

**Verdict:** Weakest of the four pages. Doesn't earn its position as the front door.

**The strategic challenge:** Roughly 80% of sales are to traders, but the remaining 20% (CAD users, financial controllers, security operations, accountants, general multi-screen professionals) is real revenue that can't be alienated. The homepage needs to confidently serve both audiences in the first 3-5 seconds, then route them to deeply specialised pages from there.

**The pattern: two-door homepage.** A clear primary identity (multi-screen specialists since 2008) with an obvious, fast route into the trader-specific journey. Traders feel immediately recognised and routed; non-trader multi-screen buyers don't feel like they've landed on the wrong site.

**Issues with current homepage:**
- "Multi-Screen Computers, Stands & Monitors" headline is descriptive rather than persuasive, and doesn't acknowledge trader specialisation
- Layout feels like a 2018 template
- No clear routing for either audience above the fold
- A trader has to scan past generic copy, blog teasers, and a "welcome to MultipleMonitors" section before they figure out where to go — non-traders have the same problem in reverse
- No Trustpilot strip, no prominent BBC mention, no "trusted by" customer logos
- Three testimonials are tucked low and are weak (see 3.5)
- Blog teaser takes up real estate that should work harder

If a Google search for "trading computer UK" sends someone to the homepage rather than `/trading-computers/`, you're probably losing them today.

**Recommended structure (top to bottom):**

**1. Hero zone — establish broad identity in one line**

Headline that acknowledges both the broad offering and trader specialisation without picking sides too hard. Working examples (refine to taste):
- "Multi-screen specialists — including the UK's leading trading computers"
- "Multi-screen computers, stands and monitor arrays — built and supported in the UK since 2008"

The first option is bolder and more honest about where the expertise lies. Recommended.

Sub-headline reinforces credibility: "Trusted by hedge funds, prop trading firms, financial institutions and thousands of professionals who need more screens."

This zone doesn't try to convert anyone yet — it just establishes that whoever you are, you're in the right place.

**2. Trust strip — immediately under hero, audience-neutral**

Single horizontal band with:
- BBC "As seen on" badge (upgraded from current tiny banner — see 3.4)
- Live Trustpilot widget with star rating and review count
- "Established 2008" / "17+ years" credential
- Optional: "1000s of trading PCs delivered" or similar volume signal

**3. Audience routing — the "two doors"**

Three-card layout giving each major segment a clear door. Visually equal weight, but the trader card carries an extra signal of specialisation:

- **"I'm a trader"** → `/trading-computers/`
  Card copy: "Specialist trading PCs with our own benchmark data. Trusted by hedge funds and individual traders since 2008."
  Visual specialisation signal: small "BBC featured" badge, or a snippet of one of the TraderSpec CPU charts, or "1000s of trading PCs delivered"

- **"I need multi-screen computers"** → `/computers/`
  Card copy: "Multi-monitor capable PCs for any professional workload — CAD, finance, security, analytics."

- **"I just need stands or monitors"** → `/stands/` and `/display-systems/`
  Card copy: "Multi-screen stands and monitor arrays, ready for any workspace."

The trader card doesn't need to be visually dominant — making it equal in size but giving it the extra specialisation signal achieves the routing without dissonance against the broader headline.

**4. Audience-neutral proof — works for everyone**

- "What makes us different" pillars: UK-built, in-house testing, 5-year hardware cover, lifetime support
- Customer logo strip (mix of trader firms and other multi-screen customers)
- Live Trustpilot reviews carousel (not the static three quotes currently used)

**5. Bundle savings callout**

Horizontal strip: "Buy a complete bundle, save £200+ vs ordering separately." Works for both trader and non-trader audiences since bundle discounts apply across the board.

**6. Brief content blocks that hint at depth**

- Teaser for TraderSpec data → links to `/trading-computers/`
- Teaser for the buyers' guide → lead capture
- Latest blog posts (2-3, not the current treatment)

These give people who want to dig in a clear next step without forcing them through a sales funnel.

**Discipline to maintain when building:**

Resist the urge to over-explain on the homepage. A common mistake on multi-audience sites is trying to address every segment's questions on the homepage itself. Don't. The homepage's job is to **recognise the visitor and route them**. The `/trading-computers/`, `/computers/`, and `/stands/` pages do the actual selling. If this discipline holds, the homepage stays clean and fast.

### 4.2 `/trading-computers/` landing page — "Trader Destination"

**Verdict:** Your strongest page today. Data-led narrative is the right approach. But its job is changing significantly.

**Strategic re-framing:** With the new homepage routing structure, this page becomes the **"trader homepage"** — the destination for everyone routed in from the main homepage, from Google ads, and from organic search on trader-specific terms. It is no longer competing with the homepage for the trader's attention; it is the destination. This means it should be deeper, longer, and more deliberate than it is today.

**Issues with current page:**
- Hero ("Specially Crafted Computers Designed To Power Your Trading Sessions With Ease") is generic and doesn't differentiate
- "Jump straight to the computers" javascript link is a clunky pattern (and `javascript:` links can also be flagged by some security/scanner tools)
- Two PC product cards at the bottom are well-presented but no comparison table — most customers comparing two products want a side-by-side
- Bundle grid below feels visually busy
- No FAQ section
- No "still not sure? talk to Darren" enquiry form
- No exit-intent capture for the buyer's guide
- Doesn't yet carry the "we are *the* UK trading computer specialists" weight that its new role requires

**Recommended changes:**

- **Proof-led hero:** "We're the UK's only trading computer specialist with our own published benchmark data. Here's what 17 years of building for traders has taught us." → leads straight into the CPU charts and TraderSpec data
- **Trader-specific trust strip** under the hero: BBC mention, Trustpilot stars, "trusted by [trader firm logos]", "1000s of trading PCs delivered since 2008"
- **Trader PC vs Trader Pro side-by-side comparison table** with star ratings tied to platform/workload
- **Decision tool / quiz** (covered in section 5.2) embedded inline rather than just listed in the action plan — this is the trader's primary "help me choose" mechanism
- **FAQ accordion** using content lifted from the long-form `/pages/trading-computers/` page's "Common Customer Questions" section
- **Bundle section** simplified visually, with the £200+ savings story made explicit (see 5.1)
- **"Not sure which one? Book a 15-min call with Darren"** inline CTA at the decision point
- **Replace `javascript:` jump links** with anchor links (`#computers`)
- **Trader-specific testimonials** — pull richer Trustpilot reviews from actual traders, ideally with platform mentioned (MT4, NinjaTrader, TradeStation, Bloomberg, etc.)
- **Sticky configure CTA** on desktop for users scrolling deep into the content
- **Lead capture for the buyers' guide** inline mid-page and as exit-intent

**On length and depth:** Don't be afraid to make this page significantly longer than it is today. A trader spending £1,500–£3,000 wants to feel they have all the information. The current page is short partly because it's competing for attention with the homepage above it; in the new structure, it can afford to be the comprehensive resource it should be.

**On the long-form `/pages/trading-computers/` page:** With `/trading-computers/` doing more of the heavy lifting, the long-form page becomes a deeper-dive option for the most considered buyers. Keep it (with the updates noted in 4.4) but think of it as a secondary asset — linked from `/trading-computers/` for "want the full story?" rather than being a primary funnel page in its own right.

### 4.3 `/computers/` page — Non-Trader Destination

**Verdict:** By far the most underdeveloped. Feels like a placeholder.

**Strategic framing:** With `/trading-computers/` serving trader audiences (~80% of sales) and carrying trader-specific specialisation, `/computers/` becomes the front door for the remaining non-trader audience (CAD users, financial controllers, security operations, analytics professionals, general multi-screen buyers). It should not mirror `/trading-computers/` — those audiences have different decision criteria that don't map to a platform-performance comparison tool.

**Product line decision (confirmed):** Keep all four PCs (Ultra, Extreme, Trader, Trader Pro) as separate SKUs. Ultra and Trader PC share a base platform with different configurations and marketing; same for Extreme and Trader Pro. The trader-specific branding genuinely helps trader customers, and non-trader buyers shouldn't be sold a machine called "Trader." Cross-page honesty about the underlying shared base is a trust-builder (see 4.3.8 below).

**Issues with current page:**
- Two product cards, a small "trading computers" callout, generic "which PC is right for you?" link, and that's about it
- No social proof, no benefits list, no testimonials, no Trustpilot, no rich content
- No clear routing for traders who land here by mistake
- Given Google Ads might land non-trading-specific traffic here, this page should be doing 5x more work

#### 4.3.1 Hero — general-purpose multi-screen positioning

Headline acknowledges the broader audience without losing specialisation credibility:

*"Multi-screen PCs for professionals who need more than two displays."*

Sub-headline bridges to existing strength:

*"Built in the UK since 2008. Same build quality, stress-testing, and lifetime support as our specialist trading computers — configured for CAD, finance, analytics, security operations, and any workload that outgrew a standard desktop."*

This gives non-traders a clear "this is for you" moment and borrows credibility from the trader specialism without making it the headline.

#### 4.3.2 Trader redirect strip — above the fold

Prominent horizontal band immediately under the hero:

*"Building a trading setup? Our Trader PC and Trader Pro are spec'd specifically for MT4, NinjaTrader, TradeStation, Bloomberg and other trading platforms. See the dedicated trading computers page →"*

A trader landing here by mistake should be one click from the right page. Do not bury this — the worst outcome is a trader buying an Ultra when a Trader PC would have been the better purchase.

#### 4.3.3 "Which machine do I need?" — by use case, not spec

**This is the key structural difference from `/trading-computers/`.** Instead of a spec comparison, lead with use cases. A grid of four or five cards covering the main non-trader buyer types:

- **CAD and engineering** — Solidworks, AutoCAD, Fusion 360, Revit
- **Finance and accounting** — heavy Excel, multiple browser platforms, accounting software
- **Security and operations** — monitoring dashboards, video walls, SOC/NOC use
- **Analytics and data work** — Power BI, Tableau, heavy spreadsheet work
- **General multi-screen productivity** — developers, writers, project managers with many windows

Each card briefly describes the workload and recommends either the Ultra or the Extreme. The recommendation does the hard work for most visitors — they don't need a configurator because they aren't comparing two products side-by-side, they're trying to figure out if this site is for them.

This pattern also scales well for a solo operator: adding a new use case is adding a card, not re-architecting a tool.

#### 4.3.4 The two PCs — clear positioning, minimal duplication

After the use-case grid, present Ultra and Extreme as the two featured product options for the non-trader audience. Keep this section tight:
- Short description of each
- Starting price
- Key spec summary
- "Who it's for" single-line positioning
- "Configure your [Ultra/Extreme]" CTA

**Do not build a full comparison table here.** The trading-computers comparison does a specific job (platform performance rating) that doesn't translate to generic workloads. A simple side-by-side "Ultra vs Extreme" summary with three or four differentiating rows (CPU class, max RAM, max screens, starting price) is enough.

#### 4.3.5 "What you get" section — shared across the line

Celebrate what all four machines share rather than trying to differentiate them:
- UK assembly, same workshop, same tech
- 32-hour stress test on every machine
- 5-year hardware cover, lifetime phone support
- Same component quality standards across the whole range
- Same graphics configurations for multi-screen support

This turns the four-SKU situation into an advantage rather than a complication.

#### 4.3.6 Trust signals — reuse homepage / trading-computers components

Trustpilot widget, BBC mention, customer logos, testimonials. Reuse the same components from the homepage and trading-computers pages — don't build bespoke versions. Operationally simpler, visually consistent.

#### 4.3.7 FAQ — workload-specific, not platform-specific

Different questions than the trader FAQ. Worth covering:
- "What's the difference between your Ultra and your Trader PC?" (see 4.3.8)
- "Will a multi-screen PC work with my existing monitors?"
- "I use [CAD / finance software / etc.] — will this handle it?"
- "How many screens do I actually need for [workload]?"
- "Do I need a workstation-class GPU?"
- "Can you spec a machine for [specific industry]?"

#### 4.3.8 Honest section on Ultra vs Trader PC / Extreme vs Trader Pro

**Keep this. It's a differentiator.** Most sites paper over product-line overlap with marketing language; being direct builds trust and handles the objection from smart buyers who've already worked it out.

Suggested copy pattern:

*"You might see a Trader PC and wonder how it relates to the Ultra — honest answer, they share the same base platform. The Trader PC is configured, spec'd and tested specifically for trading software workloads, with benchmark-tuned options. The Ultra is the same base machine with a broader range of options for general-purpose multi-screen work. Same build quality, same workshop, same 5-year cover. Pick the one whose configuration and positioning fits your use case — the underlying hardware quality is identical. Same story for the Extreme and Trader Pro."*

This kind of transparency is unusual. It makes the business look confident and honest, not confused.

#### 4.3.9 Cross-link to trading computers at the end

Final section acknowledges traders again: *"Trading? Our Trader PC and Trader Pro are configured specifically for trading platforms, with published benchmark data — see the full trader destination."*

#### 4.3.10 What NOT to build on this page

A few temptations worth resisting:
- **A four-way configurator.** Four PCs side-by-side in a comparison tool is unworkable. Don't try. Non-traders aren't making the same two-way comparison traders are.
- **Separate comparison tools for each use case.** Overengineered. The use-case grid with recommendations does this job in a much simpler way.
- **Deep technical content rivalling `/trading-computers/`.** The TraderSpec advantage is specifically for trading workloads. Producing equivalent benchmark depth for CAD, Excel, or security software would be huge effort for a much smaller audience. Accept that `/trading-computers/` is the technical flagship; `/computers/` is the approachable front door.

### 4.4 `/pages/trading-computers/` long-form page

**Verdict:** Excellent in places but showing its age.

What works:
- Story-telling and competitor-debunking sections are great
- "Common Customer Questions" content is gold

Issues:
- Prices are out of date
- References to Falcon and Digital Tigers are US-centric (newer UK readers won't know who they are)
- Screenshots are old and low-res (forum screenshot, test results image)
- One giant scroll with no in-page navigation, no expandable sections, no sticky CTA
- Doesn't link out to live Trustpilot reviews

**Recommended changes:**
- Update all prices and remove "brand new" claims
- Replace US competitor references with relevant UK comparison points
- Add a sticky table of contents on desktop
- Add a sticky "Configure your Trader PC →" CTA bar
- Refresh screenshots in higher resolution
- Repurpose the "Common Customer Questions" section as a structured FAQ on `/trading-computers/` where more people will actually see it

### 4.5 `/bundles/` page — Surface the Builder, Don't Rebuild It

**Verdict:** The underlying bundle builder is genuinely good. Customer praise is real and the mechanics work. The problem is purely presentational — the page is structured in a way that makes the builder feel like a secondary feature when it should feel like the main event.

**What's working (protect this):**
- Three-step flow (stand → screens → PC) is sensible and well-liked
- Auto-selection of appropriate graphics option based on stand screen count is exactly the kind of invisible-but-important detail that makes a configurator feel polished
- Final "PC product page with bundle summary header" approach gives customers full product detail plus bundle context
- URL parameter pattern (`sid`/`mid`/`cid`) preserves shareable bundle state — see 4.5.8 below
- Copyright year is already dynamic (2026 showing correctly)

**The core problem:** First-time visitors see six example bundle cards in standard product-card format and reasonably conclude they've seen the complete product line. A very normal scrolling behaviour stops before the "Create Your Own Bundle Deal" section, which is the feature customers actually love. The "scroll down" link in the subheader is the only signal that a builder exists, and subheader text is one of the most routinely-ignored areas of any page.

**This is a layout problem, not an architecture problem.** Solid afternoon-to-one-day refactor, not a major project.

#### 4.5.1 Flip the page order — the single highest-impact fix

Put the builder at the top as the hero proposition; move example bundles to the bottom as "not sure where to start? here are some popular configurations."

Suggested structure:

1. Hero — positioning bundle offer and savings (keep most of what's there)
2. **"Build your bundle"** — visually dominant section that makes the three-step builder feel like the main event
3. **"Why a bundle?"** — savings breakdown, free extras, free delivery (see 4.5.4)
4. **"Popular starting points"** — the six example bundles, positioned as inspiration or quick-starts rather than *the* product line
5. Trust signals, FAQ, etc.

This one change probably captures 80% of the value of any improvement to this page. The builder goes from hidden feature to headline; example bundles become what they actually are — shortcuts for people who don't want to build from scratch.

#### 4.5.2 Make the builder hero section feel like the main event

Current stand selection is functional but visually flat (twelve stand cards in a grid with "Select This" buttons). Works, but doesn't feel like a main-event section.

**Recommended builder hero pattern:**

- **Visible three-step flow** — "1. Choose your stand → 2. Pick your screens → 3. Select your PC" — shown as numbered steps across the top, step 1 highlighted. Makes the journey visible and signals that the user *can* do all three. Even if they don't engage immediately, they now know the builder exists.
- **Featured stand categories** rather than the full twelve-stand list. Three or four grouped cards covering typical configurations: "Dual screens," "Triple / Quad," "Six screens," "Eight screens." Clicking any takes them to step 2 (or expands inline to show specific stand options for that category). Reduces decision overwhelm — twelve stand options upfront is a lot when you're still trying to understand what a bundle is.
- **A "see all stands" link** for buyers who know exactly what they want.
- **Running bundle preview** visible alongside the builder. As selections are made (or even before), show "Your bundle so far: [stand] + [screens] + [PC] = £X total, £Y saved." Even if it starts as an empty/ghost version, it demonstrates the end state and makes the savings tangible.

#### 4.5.3 Show what a finished bundle actually looks like

Ties directly to the photography discussion (Wave 1 of the world-class work). A finished six-screen bundle is much more impressive in a real photo than as an icon-stack of stand + screens + PC. Currently the page doesn't show the end-result aesthetically.

Three photography opportunities, in priority order:

1. **Lifestyle hero image** showing a real six-screen bundle on a real desk, running actual trading platforms (composite approach). Replaces or sits alongside the current `bundles.png` graphic. Does more work than any amount of copy.
2. **Bundle cards with real photos** instead of grey product-shot style. Each example bundle shown as a real photo of that configuration in use. Becomes part of the decision support.
3. **"As delivered" unboxing visual.** One of the under-sold bits of the bundle proposition is that everything arrives in one delivery with all cables. A photo of a bundle as it arrives (boxes stacked, cables bagged, PC on top) visualises the hassle-free promise in a way bullet points can't.

The lifestyle hero alone would significantly change how the page feels.

#### 4.5.4 Replace the bulleted savings list with concrete example comparison

Current copy: *"Up to £100 in discounts / Free Speakers (Save £20) / Free Wifi Card (Save £40) / Free UK Delivery (Save £20) / Free Long Length Cables / Total Savings Of Up To £180!"*

Bulleted components feel theoretical. On the trading-computers mockup the framing is stronger: *"£214 avg. saving on a 6-screen bundle vs piecing it together."* Apply that here — show what a real customer saves on a real example bundle, not a list of components with individual values.

A before/after price comparison ("six-screen bundle: £2,100 separately, £1,900 as a bundle, plus £180 of extras included") is worth more than six savings bullets.

#### 4.5.5 Mention that all stands work with all computers

**All monitor arrays of any size are compatible with any of the four computers.** This is worth saying out loud somewhere on the builder page — probably as a short reassurance line near the stand selection step, something like: *"All our stands work with any of our computers — pick the stand you want first, we'll handle the graphics spec on the PC side."*

This removes a subtle hesitation ("am I going to pick a stand and then discover it only works with one specific PC?") and signals that the customer can make each decision independently without painting themselves into a corner.

#### 4.5.6 Small copy and CTA refinements

- "View Bundle Details" is weirdly formal — "Configure this bundle" or "See this bundle" is more direct
- Example bundle captions should position them as starting points: "A popular 4-screen starting point — customise anything before checkout" rather than just "A Quad 24" Pyramid Monitor Array and an Extreme Multi-Screen Computer"
- "Create Your Own Bundle Deal" heading could be punchier — "Build your own bundle" or "Configure your setup"

#### 4.5.7 Consider surfacing the "save/share this bundle" capability

The URL parameter pattern (`sid`/`mid`/`cid`) means every configured bundle already has a shareable URL — customers can bookmark a specific configuration, email it to a business partner, or come back to it later. This is genuinely useful for high-value purchases where buyers often discuss with a partner before committing, but the capability isn't currently surfaced as a feature.

A simple "copy link to this bundle" button on the final product page would expose an existing capability without any backend work. Low effort, clear value for buyers who need to think it over or get sign-off.

#### 4.5.8 What NOT to build on this page

Temptations worth resisting:

- **Drag-and-drop bundle canvas where stands/screens/PCs are pieces on a board.** Complex, fragile, and the current three-step linear flow is already working. Customer praise confirms it.
- **A "suggest a bundle based on my use case" quiz.** Resist copying patterns from `/trading-computers/`. Bundle buyers usually already know how many screens they want. A quiz adds friction to solve a problem customers don't actually have.
- **Cross-sell additions during the builder flow** (UPS? keyboard? second stand?). Increases cart friction without materially increasing order value. Keep the bundle flow clean — accessories can be cross-sold at cart or post-purchase.
- **"X bundles built this week" / "Y people viewing now" social-proof widgets.** Tacky, and competent visitors see through them. The trust-building work you're doing elsewhere (Trustpilot, testimonials, BBC) is the right kind.

#### 4.5.9 Summary — scope and effort

Not a major project. Specifically:

1. Reorder the page: builder at top, examples at bottom
2. Restructure the builder section: numbered steps, grouped stand categories, running preview
3. Replace bulleted savings list with concrete before/after comparison
4. Add compatibility reassurance line ("all stands work with all computers")
5. Add lifestyle hero photo when photography is done (Wave 1 work)
6. Update example bundle captions and CTA language
7. Expose the shareable-bundle-URL capability via a copy-link button

None of these require backend changes. The ProductCart flow, the sid/mid/cid URL structure, and the auto-graphics-selection logic all stay exactly as they are. This is purely restructuring the entrance to an already-working funnel.

### 4.6 `/display-systems/` and the Final Array Page

Two related pages with different problems. The main `/display-systems/` page has the same "builder is hidden" issue as `/bundles/` and takes the same treatment. The final array configuration page (e.g. `/display-systems-3/?sid=312&mid=317`) is a bigger opportunity — it's the page the customer sees at the moment of commitment for a £300-£900+ purchase and currently undersells itself significantly.

**Housekeeping item first:** The final array page footer still shows "©Copyright 2022" while the main display-systems page has been updated to 2026. Worth a template sweep to find any other pages with the old year.

#### 4.6.1 Main `/display-systems/` page — apply the bundles treatment

Essentially the same structural problem as `/bundles/` (see 4.5):
- Six example arrays at the top look like the complete product line
- The "Create Your Own Monitor Array" builder is hidden below them
- "Scroll down to create your own" signal lives in a subheader that most visitors skip

**Fixes transfer directly from 4.5:**
1. Reorder: builder above, examples below
2. Restructure builder with numbered steps (stand → screens) and grouped stand categories
3. Add running preview of the configured array as selections are made
4. Replace the "Array Bonus" bulleted list with a concrete savings example
5. Update example array captions to position them as starting points rather than *the* product line
6. Lifestyle photography (Wave 1)

**One array-specific adjustment:** The current array offer ("Free Long Length Cables Worth £15 Per Screen / Reduced Delivery Fee") is genuinely weaker than the bundle offer, because arrays don't include the PC so the wifi/speakers/bundle discount don't apply. Don't oversell this — instead, reframe the value proposition away from savings and toward the "matched, aligned, tested-together" story. The real value of buying an array vs buying a stand and screens separately isn't primarily cable bonuses — it's that someone else has checked compatibility, matched bezel widths, verified VESA specs, and packaged it as one delivery. That's worth more to the customer than the cable bonus, and it's currently invisible.

Also: "Reduced Delivery Fee" is vague. The final array page shows "With Free Delivery (Save £10!)" on at least some configurations. If delivery is free above a threshold, state it concretely — "Free UK delivery on arrays over £X" is more persuasive than "reduced delivery fee."

#### 4.6.2 Final array page — where the real opportunity is

This is the page the customer sees after making two decisions (stand, screens). They're primed to buy. But the page reads as a receipt, not a purchase decision. For a £500+ purchase from a business the customer may not have bought from before, that's a large gap.

**Current state:** One small hero image, three text blocks summarising the stand / screens / cables, a price, a buy button. No trust signals, no testimonials, no specifications beyond model names, no photos of the array in use, no warranty or delivery reassurance at point of commitment.

Structural recommendations below.

#### 4.6.3 Show what they're actually buying

The single most important change. A customer buying a triple 24" array wants to see:
- Three 24" screens actually mounted on the horizontal Synergy stand
- The array on a desk with scale reference
- Close-up of the mount mechanism
- Back of the setup showing cable routing
- What's in the box when it arrives

Ties into the Wave 1 photography investment. Arrays should be included in the same photography session as the trading computer setups — same composite approach (real hardware, real screens, real content), same session.

#### 4.6.4 Full specifications, not just model names

Current page tells the customer they're getting "Three Acer 24" 1920 x 1080 (Full HD) widescreen monitors" and that's it. A £500+ buyer at decision point wants:
- Exact model of the monitors
- Panel type (IPS, TN, VA)
- Refresh rate
- Inputs (HDMI, DisplayPort, USB-C)
- Bezel width (critical for multi-screen setups)
- VESA compatibility
- Dimensions and weight of the full assembly
- What's in the box (VESA plates, screws, tools, cable ties)
- Warranty on screens vs stand

ProductCart may limit how much underlying product data you can surface here. At minimum, add a "Full specifications" expandable section covering the above. If you can pull the underlying product details dynamically from the stand and monitor records, the page immediately becomes significantly more useful for the serious buyer.

#### 4.6.5 Add trust signals at the moment of commitment

Currently zero on this page. Add:
- Trustpilot star rating and review count
- Lead time ("Ships in 2-3 working days")
- Warranty summary (stand: X years, screens: Y years)
- Return policy link / reassurance
- "Free UK delivery" made prominent
- One relevant customer testimonial

These matter much more on the final page than on the landing page. This is where reassurance converts.

#### 4.6.6 Take credit for the compatibility-checking work

Implicit but never stated: when a customer buys a specific array, they're trusting that you've checked the screens fit the stand, the mount hardware is compatible, the VESA specs line up, the bezels match, and it works when it arrives. That's real value — take credit for it.

Suggested copy pattern:

*"We've checked: the Triple Horizontal Synergy Stand is fully compatible with these Acer 24" monitors. Matched bezel widths, matched VESA mounting, and they've been physically test-fitted in our workshop. Arrives as one delivery with every mount plate, screw and cable you need."*

Both reassures the anxious buyer and positions the "configured together" element as a feature rather than an invisible given.

#### 4.6.7 Add the PC / bundle cross-link

A customer configuring an array either already has a compatible PC, or is thinking about getting one. If the latter, surfacing the bundle option here prevents them from buying the array separately and then separately buying a PC — worse for them (no bundle discount) and worse for order value.

Suggested callout:

*"Need a computer to drive this array? Add a Trader PC or Ultra PC as a bundle and save £X — free cables, free wifi, free delivery. [See bundle options →]"*

Not pushy — factual, with a clear value proposition.

#### 4.6.8 Handle the "is this the right array for me?" lingering question

Not every customer reaching the final page is 100% committed. They may still be comparing this config to a quad or pyramid. Low-effort safety nets:
- "Wanted a different size or layout?" link back to main page
- "Looking for something we don't show? Call Darren on 0330 223 66 55"
- Small "customers who bought this also considered" section with two or three alternative configurations

Not heavy cross-selling — just a safety net.

#### 4.6.9 Refine the "Change Selection" links

Currently two "Change Selection" links (one for stand, one for screens) that send the customer back up the funnel. Two improvements:
- **Rename.** "Change stand" and "Change screens" reads more naturally than "Change Selection."
- **Inline selection (future iteration).** Ideally these become dropdown or modal selectors on this page so the customer can iterate without restarting. Larger technical change — worth considering for a future iteration but not essential.

#### 4.6.10 Expose the save/share capability

Same point as bundles (4.5.7). The URL parameters (`sid`/`mid`) mean every array has a shareable link. A "Copy link to this array" button surfaces the capability for buyers who need to discuss with a partner before committing. Trivial to add, genuinely useful.

#### 4.6.11 URL structure and SEO note

Cosmetic / future consideration: `/display-systems-3/` exposes internal funnel structure in the customer-facing URL. If ProductCart supports it, more human-readable URLs (`/array/triple-24-acer/` or similar) would help sharing and SEO. Also worth checking that each dynamically-generated array page has unique titles and meta descriptions — currently the same page title ("Triple 24" Monitor Array") would appear for both the Acer and Iiyama versions of the same physical configuration. Out of scope for immediate improvement pass but worth flagging for future SEO work.

#### 4.6.12 What NOT to build

- **Don't replace the final array page with a full product-listing-style page.** Customer has already made their choices — don't restart them.
- **Don't add per-array reviews.** Spreading Trustpilot reviews across dozens of configurations makes individual pages look empty. Use aggregate Trustpilot rating instead.
- **Don't add pushy upsells** (cable upgrades, accessories). Customer is at commit moment — adding friction here is costly. The PC/bundle cross-link (4.6.7) is the one exception because it's genuinely relevant to the purchase.
- **Don't build a full configurator-within-configurator.** The inline-selection idea in 4.6.9 is optional and low-priority.

#### 4.6.13 Summary — scope

**Main `/display-systems/` page:** one afternoon. Apply 4.5 bundles fixes.

**Final array page:** 2-3 days once photography is in hand. Meaningful work is bringing in real imagery, surfacing screen specifications, adding the trust signal block, adding the compatibility-checked reassurance copy, and adding the PC/bundle cross-link. Photography can wait for Wave 1; everything else can be done now.

### 4.7 `/stands/` page — Tell the Synergy Stand Story

**Verdict:** Significant missed opportunity. The page presents a genuinely unique, proprietary, UK-manufactured modular stand system as a commodity list of twelve products with thumbnails and prices.

**What's going wrong:** Nothing on the current page tells the visitor that these stands are unique to Multiple Monitors, UK-designed and UK-manufactured, a modular system, proprietary to this business, or that they've been refined over 17 years with proven real-world demand (200-500 units sold annually). The single line "Strong | Modular | Stable | Flexible | Attractive" with a small "Discover Why" link is the only gesture toward the provenance story, and most visitors will never see it.

**The commercial implication:** At 200-500 units per year at £115-£375 each, this represents meaningful revenue (~£40k-£150k) with good margins. The current page treatment is leaving money on the table by failing to justify premium pricing against generic Amazon/VIVO competitors.

**The strategic shift:** The stands page needs to stop being a product grid and start being a brand story page that happens to have products on it. Done well, a visitor should feel they've discovered something better than generic mounts, and that £175 for a triple stand is a fair price for what they're actually getting.

#### 4.7.1 Hero — lead with the brand story

Current hero: *"HOLD YOUR SCREENS IN PERFECT ALIGNMENT WITH A STRONG AND EASY TO SETUP MULTIPLE MONITOR"* (truncated mid-sentence, generic, says nothing unique).

Replacement pattern:

**Headline:** *"Synergy Stands — our own UK-designed, UK-manufactured modular monitor mounts."*

**Sub-headline:** *"Developed by us, manufactured in the UK to our specifications. A modular system that scales from two screens to six on a single assembly. Built to hold up day after day, sold and delivered since 2008, with thousands in use across trader desks, design studios and operations rooms."*

Primary CTA: "See the range" (anchor to product grid). Secondary: "How the modular system works" (link to existing `/pages/synergy-stand/`).

#### 4.7.2 "What makes a Synergy Stand different" — benefit cards

Missing entirely from current page. A section sitting between hero and product grid with 4-5 distinct benefit cards:

- **UK-designed and UK-manufactured.** Most competitors sell rebranded imported generic mounts. You don't. Say it clearly — particularly valuable for business buyers who care about provenance.
- **Modular system.** Start with two screens, add to four, extend to six. Show this visually — photo or diagram of the same base stand in different configurations.
- **Built for daily professional use.** Not gaming stands, not home office stands. Made for setups running all day, every day, for years. 17 years of iteration and 200-500 units sold annually is the proof.
- **Adjustability that actually works.** Describe what the stand does in use — height, tilt, rotation ranges — not just that it's "adjustable."
- **Single-vendor compatibility.** Every Synergy Stand works with every monitor sold on the site. No mount-plate surprises, no "does this fit?" anxiety. This is a real anxiety that drives buyers to Amazon — address it directly.

#### 4.7.3 The design and manufacturing story — honest framing

**Credibility handling:** A direct "I personally designed these" claim carries risk if challenged and doesn't carry the weight it should. Hiding the provenance entirely (current page) is worse. The middle path is stronger than both — position this as a company achievement with specialist partners, not a one-person engineering claim.

Suggested copy pattern:

*"The Synergy Stand started in 2008 as a solution to a problem we kept hearing from customers: the generic stands they could buy elsewhere weren't strong enough, stable enough, or flexible enough for real multi-screen work. After years of supplying customers with what the market offered, we decided to build what the market didn't.*

*Working with a specialist UK design and manufacturing partner, we developed the Synergy Stand system from scratch — designed around the specific needs of multi-screen users, tested by us, refined through multiple generations based on real customer feedback. Every stand we ship is manufactured in the UK to our specifications and packaged in our workshop."*

Further strengthening options:
- Name the manufacturer if comfortable doing so (some customers appreciate the transparency)
- Reference specific generations / revisions if they're real ("Now in its Xth generation")
- Personal framing that doesn't overclaim: *"Darren Atkinson founded Multiple Monitors in 2008 after years of frustration with the multi-screen mounts available at the time. The Synergy Stand range is the result of a decade-plus collaboration with a specialist UK design and manufacturing team to produce the stands he knew the market needed."* — positions founder as the vision, not the engineer. Both more truthful and more credible.

#### 4.7.4 Photography — replace product thumbnails with real imagery

Current thumbnails are technical product shots against white backgrounds with no monitors attached. They communicate nothing about the stand in use.

Photography priorities (part of Wave 1 session):
- Hero photo of a real six-screen Synergy Stand fully assembled on a real desk, screens showing real content (composite approach)
- Detail shots of mount mechanisms, adjustment points, build quality
- Assembly photo showing modular components laid out
- Scale reference (with person, chair, or desk at known size)
- Workshop photos of stands being assembled or packaged

On product cards: keep the current technical shots but consider a secondary "view in use" image showing the stand with real screens.

#### 4.7.5 Restructure the product grid by screen count

Current: twelve stands in a flat grid with no grouping. Buyers have to read every caption to figure out what they're looking at.

Regroup with section headers:
- **Dual-screen stands** (2 products — vertical and horizontal)
- **Triple-screen stands** (2 products — horizontal and pyramid)
- **Quad-screen stands** (4 products — square, pole, pyramid, horizontal)
- **Five-screen stands** (1 product)
- **Six-screen stands** (2 products — pole and side-by-side)
- **Eight-screen stands** (1 product, noting it's two quad-square stands in a 2-over-2 configuration)

Reduces cognitive load significantly. Buyer can skip straight to their screen count.

#### 4.7.6 Show the modular upgrade path

Modularity is one of the strongest differentiators and currently invisible. Dedicated section — ideally visual:

*"Starting with 2 screens? The same base stand accepts additional arms. Scale up to 4, 5 or 6 screens as your needs grow — no need to buy a whole new stand."*

A simple "2 → 4 → 6" visual schematic, or photographs of the same stand in three configurations, communicates this better than any paragraph. Matters for:
- Buyers uncertain about future screen count (removes decision friction)
- Value perception (you're buying into a system, not a fixed product)

#### 4.7.7 Technical specifications

Most buyers at £115-£375 want:
- Max screen size supported (per-screen and total)
- Max weight per monitor
- VESA compatibility (75x75, 100x100, both)
- Materials (steel, aluminium)
- Base footprint / desk space required
- Height adjustment range
- Tilt / rotate range
- Weight of the stand itself
- Assembly time and whether tools are included
- Warranty length

A "Technical specifications" expandable section or a comparison table across all stands. Since these are designed to your specifications, you have this data more readily than any competitor — expose it as a differentiator.

#### 4.7.8 Trust signals

Currently zero on this page. Add:
- Trustpilot rating and review count
- "Sold since 2008, thousands in use"
- Warranty summary
- UK delivery promise
- Stand-specific testimonials if available (more persuasive than generic)

#### 4.7.9 Improved bundle cross-link

Current "Save Money, Get Free Cables & Free Delivery with a Bundle" callout is small and buried at the bottom. Bundle cross-sell is particularly relevant here — a customer buying a six-screen stand probably also needs six screens. Reframe:

*"Buying a stand? Save £200+ when you add screens, a PC, and all the cables as a complete bundle. Free cables, free wifi, free UK delivery. [See bundle options →]"*

#### 4.7.10 Consolidate the `/pages/synergy-stand/` content and retire the URL

The content on `/pages/synergy-stand/` is stronger than expected after a direct read — it contains four genuinely useful sections that should be ported to `/stands/`, after which the old URL should be 301-redirected to `/stands/`. The existing page is a dead-end (no configurator, no pricing, no product grid) and is reached by almost no visitors, so it's doing close to zero conversion work despite containing good material.

**Four pieces of content worth porting to `/stands/`:**

1. **All-steel vs plastic argument.** Currently on the synergy-stand page as "Strength & Stability." One of the sharpest pieces of copy on the site — calls out competitors directly ("lots of other stands try to save costs by using a metal frame combined with a range of plastic parts") and justifies the price premium. Becomes one of the "What makes a Synergy Stand different" benefit cards (4.7.2). Tighten from four paragraphs to two.

2. **28" screen capability and curve argument.** Currently on the synergy-stand page as "Supports Larger Screens." The specific insight that competitors' "up to 24"" claims leave no room to curve the outer screens is particularly good — buyer-relevant information that addresses a real problem. Becomes another benefit card. Keep the competitor framing.

3. **Modular explanation.** Currently on the synergy-stand page as "Modular Build System." The "voice in your head saying wouldn't it be easier if I just had one more screen" framing is in-character for the brand voice. Forms the basis of the modular upgrade path section (4.7.6).

4. **Flexibility / adjustability detail.** Currently on the synergy-stand page as "Flexible Enough For Any Requirement." Specification-grade detail about height adjustment, hinge mechanism, horizontal slide, pivot, tilt, 30mm fine adjustment. Belongs prominently on the main `/stands/` page as a "Designed for real-world use" section — this is the kind of detail that differentiates from Amazon mounts.

Together these four pieces deliver most of the "What makes a Synergy Stand different" section the main `/stands/` page is currently missing.

**Content to drop, not port:**

- **The "dedicated Synergy Stand website" reference** in the modular section. Refers to the old Synergy Stands site that didn't work out — the URL is no longer live, so this line actively promises a capability that doesn't exist. Must be removed regardless of the consolidation decision.
- **The closing "Make the right choice" CTA** that just links back to `/stands/`. On the main page this is circular. Replace with a proper configurator or product-grid jump link.

**URL handling:**

301-redirect `/pages/synergy-stand/` to `/stands/` once content is migrated. Before redirecting, check Google Search Console for:
- Any inbound external links to the URL
- Any queries the page currently ranks for (even weakly)

If the page has any ranking authority, the redirect preserves it and passes it to `/stands/`. Deleting the page without a redirect loses that. The redirect should be in place before any content migration goes live to avoid a window where the old URL returns 404s.

**Why redirect rather than keep as deeper-dive:**

The "Inside the Synergy Stand" educational-piece option considered earlier looked less attractive once the actual content was read — the material isn't deep enough to justify standalone existence after the best parts are migrated, and keeping two pages covering the same product splits topical authority for SEO without meaningful user benefit. Consolidation is cleaner.

**On the "attractive" claim:**

The main `/stands/` page positions stands as "Strong | Modular | Stable | Flexible | Attractive" but neither page currently demonstrates the "attractive" claim — images are technical product shots against beige backgrounds. If the "attractive" positioning is kept, the Wave 1 photography session (4.7.4) needs to show stands looking good on real desks, not just functional. If the photography doesn't support the claim, drop "Attractive" from the positioning line.

#### 4.7.11 What NOT to do — dedicated Synergy Stand domain

**Historical context:** A dedicated Synergy Stands website existed previously with its own checkout, content, and positioning. It didn't get traction and conversion was better on multiplemonitors.co.uk. This is consistent with the tradingcomputers.co.uk analysis earlier — split sites generally dilute SEO authority and operational focus without compensating benefits for a business at this scale.

**The lesson to apply:** Don't try to spin the stands off again. Instead, make the stands story unmissable on the main site — hero treatment on `/stands/`, consolidated content from the synergy-stand page, and proper cross-linking from bundles and arrays pages. A strong internal story beats a weak separate site.

If the stands ever became a much larger revenue line (say, £500k+) or attracted a distinct buyer audience not overlapping with the PC buyers, the separate-site question might be worth revisiting. Not now.

#### 4.7.12 Possible future considerations (not current scope)

For reference, not for this pass:

- **Product-line naming convention.** Most stands are named descriptively ("Triple Horizontal Synergy Stand"). A cleaner convention ("Synergy Dual," "Synergy Triple Pyramid") might feel more like a product family. Minor — don't chase this unless a rebrand is already on the table.
- **Customer case-study features.** If any trading firms or hedge funds would permit a mention ("In use at [X prop firm], [Y trading desk]"), this carries real B2B weight. Gather this opportunistically over time.
- **Stand-specific photography sessions.** When Wave 1 photography happens, ensure stands get dedicated attention rather than only appearing as backgrounds for PC shots.

#### 4.7.13 What NOT to build

- **Don't add a configurator to this page.** The modular upgrade path is better served by visual/copy explanation than an interactive builder. The bundles and display-systems builders already cover the "configure a combination" need.
- **Don't try to turn this into a trader-specific stands page.** The stands serve all audiences equally — traders, CAD users, analysts, SOC operators. General-purpose product page is the right frame.
- **Don't cross-sell non-stand accessories here** (mounts for laptops, desk risers, etc.). Keep the page focused on Synergy Stands.

#### 4.7.14 Summary — scope

Two to three days of work. Product mechanics (grid, product pages, add-to-basket) unchanged. Changes are structural and content-focused.

Deliverables:
1. New hero with brand-story positioning
2. "What makes a Synergy Stand different" benefit card section
3. Design and manufacturing story with honest credibility framing
4. Product grid regrouped by screen count with section headers
5. Modular upgrade path visual/copy
6. Technical specifications section or comparison table
7. Trust signals strip
8. Improved bundle cross-link
9. Migrate the four content pieces from `/pages/synergy-stand/` (all-steel argument, 28" capability, modular explanation, flexibility/adjustability detail) and 301-redirect the old URL
10. Photography additions when Wave 1 completes — must demonstrate the "attractive" claim if it's kept in the positioning

### 4.8 Product pages — Trader PC, Trader Pro, Ultra, Extreme

**Verdict:** Four product pages with a visible consistency gap. The Trader PC and Trader Pro pages have significantly richer persuasion content (nine benefit blocks, in-configurator helper text with trader-specific framing). The Ultra and Extreme pages have identical configurators with star ratings and monitor panels working correctly, but lack the per-option helper text and the benefit blocks above. This creates a two-tier feel across the product line.

**The real issue:** A custom configurator system powers the trader PCs, a maintenance burden, while the native ProductCart system powers the others. The gap customers actually notice is the helper text under each configuration option — star ratings and monitor/resolution panels work correctly on all four.

#### 4.8.1 Role of the product pages in the customer journey

A buyer reaching a product page has typically already decided which PC they want (from the landing pages, Google ads, or as a returning visitor). Persuasion work is largely done. The product page needs to:

1. Not undo the landing-page work by feeling inconsistent or under-polished
2. Help the customer make configuration decisions (CPU, RAM, GPU options) confidently
3. Handle late-stage objections (warranty, support, delivery timelines)
4. Make it easy to convert

This framing matters because it pushes back against the temptation to re-sell the PC on the product page. The landing pages (`/trading-computers/`, `/computers/`) do the selling. Product pages close the sale.

#### 4.8.2 The configurator helper-text question

**Current state:**
- Trader PC and Trader Pro: custom configurator with trader-specific helper text under each option ("RAM dictates how many charts and programs you can hold open without slowing down your PC")
- Ultra and Extreme: native ProductCart configurator with option labels but no helper text

**The trade-off:** Audience-specific helper text converts better but is painful to maintain across two systems. Generic helper text maintains centrally but loses some sharpness.

**Four options considered:**

1. **Status quo** — keep both systems. Works if Trader PC price/option changes are infrequent (say, twice a year). Hidden cost is forgetting to keep Ultra/Extreme in sync when the trader PCs update.
2. **Generic helper text added to native system** — one set of audience-neutral helper descriptions that work for any buyer, applied to all four products. Captures ~80% of the helper-text value. Lower sharpness on trader-specific wording but single-source maintenance.
3. **Per-product helper text in native system** — extend ProductCart to support per-product (or per-product-group) helper overrides. Two sets of content (trader / non-trader), one system. More development upfront, cleaner long-term.
4. **Hybrid** — generic helper text shared across all, plus optional audience-specific callouts on products where the sharpness is worth the extra maintenance.

**Recommendation: Option 2 in the short term.**

Add simple, generic helper text to each configuration option in the native ProductCart system. Rewrite each line to work for any professional user — not trader-specific, not CAD-specific, just "here's why this option matters." Apply to Ultra and Extreme immediately.

Generic example (CPU):
*"The CPU is the biggest factor in how fast your computer feels in day-to-day use. More cores help with multi-tasking and intensive workloads."*

Compare to the existing trader-specific version:
*"The number one difference to how fast your computer will perform, CPUs impact speed and multi-tasking performance levels."*

The generic version captures the core job — tells the buyer this matters, tells them what changes between options. Most buyers of all types get enough from it.

**If per-product overrides turn out to be quick to build** (Option 3 baked into the implementation), add them. If it starts to over-complicate or delay the helper-text rollout, stick with Option 2 and ship.

**Sharper audience-specific messaging lives where it does the real work:**
- `/trading-computers/` landing page — the benchmark data, the platform comparisons, the decision tool
- The benefit blocks above the configurator on product pages (see 4.8.3)
- Product-specific FAQs (see 4.8.7)

These places justify the maintenance burden because they do more persuasion work than per-option helper text.

#### 4.8.3 Benefit blocks — trim Trader PC/Pro, add to Ultra/Extreme

**Trader PC / Trader Pro: reduce from nine benefit blocks to 3-5.**

Nine blocks is past the point where readers start skimming. By the time the new `/trading-computers/` landing page carries most of this content in richer form, the product page doesn't need to repeat the full pitch — the buyer has already decided.

Keep (consolidated into 3-5 blocks):
- CPU / platform performance (with link to comparison tool on `/trading-computers/`)
- Multi-screen support
- DisplayFusion
- 5-year cover & lifetime support
- Real customer logos / trust signals

Drop or consolidate:
- Standalone "silence" block → fold into "UK-built with care" block
- Standalone "Windows 11 tuning" block → fold into a general quality block
- Standalone "build & test" block → fold into 5-year cover / support block
- "Extra kit" block → inline this within the configurator (where mouse/keyboard/wifi already appear as options)

The goal: five strong blocks that get read beats nine decent blocks that get skimmed.

**Ultra / Extreme: add a proper benefit block section.**

Currently near-zero persuasion content between the description and configurator. Add a 3-5 block section with non-trader framing:

- Use-case framing ("Perfect for CAD and engineering, finance and analytics, security operations, and general professional multi-screen work")
- UK assembly and build quality
- Multi-screen support (same content as trader pages — this is universal)
- DisplayFusion (same — universal)
- 5-year cover & lifetime support (same — universal)

Most of this content already exists on the trader pages. Port the universal bits, swap the audience framing where needed. Don't rebuild from scratch.

#### 4.8.4 Trust signals strip — all four pages

Currently missing on all four product pages. At the price point (£925 — £1,695+) and at the moment of commitment, trust signals are critical.

Add near the price/CTA on all four pages:
- Trustpilot star rating and review count
- "As seen on BBC Traders: Millions by the Minute"
- "5-year cover · UK-built · Lifetime support" strip
- Volume credibility ("1000s delivered since 2008" or similar)

Consistent across all four product pages. Same components reused.

#### 4.8.5 Cross-links between product pages

Each product page should help the customer quickly compare to its closest alternative.

- **Trader PC page** → "Considering the Trader Pro instead?" + link. "Not sure? Compare them side-by-side on our trading computers page →"
- **Trader Pro page** → "Considering the Trader PC instead?" + link to comparison.
- **Ultra page** → "Considering the Extreme instead?" + "Trading? Our Trader PC is spec'd for trading workloads →"
- **Extreme page** → "Considering the Ultra instead?" + "Trading? Our Trader Pro is spec'd for trading workloads →"

A customer who realises mid-configure that they're on the wrong machine shouldn't have to back-navigate to figure out which is right. Cross-links reduce friction at exactly the moment when a drop-off is most likely.

#### 4.8.6 Starting CPU on the Ultra

The Ultra starts with an Intel i3 14100F (4 cores, 8 threads) at £925. This is a noticeably weak starting spec for the price — most buyers will upgrade. For comparison, the Trader PC starts at i5 14600KF for a similar price.

Two options worth considering:
1. **Lower the Ultra starting price** to reflect the i3 base (makes it genuinely entry-level)
2. **Strengthen the starting CPU** to match what most buyers would actually configure (i5 14400F or similar)

Option 2 is probably the better commercial choice — customers landing on "£925 + VAT" anchor on that number, and most then discover the practical starting spec via upgrades. An honest starting spec at a slightly higher price converts better than an underpowered starting spec that requires an upgrade to be useful.

Not a critical issue but worth reviewing when next touching Ultra pricing.

#### 4.8.7 Product-specific FAQ sections

The trader landing page FAQ covers general trader-buyer questions ("do I need a specialist trading computer?"). The product pages should have a smaller, more specific FAQ covering questions about *this particular machine*:

- "Can I upgrade the CPU later?"
- "Does this come with Windows pre-installed?"
- "What's included in the box?"
- "How do I connect my existing monitors?"
- "Can I add more storage later?"
- "What happens if a part fails?"
- "How do I move my software from my old PC?"

Pull these from actual pre-sale email inbox — not a generic "PC FAQ" template. Four to six questions per product page is about right; more becomes a wall of text in the wrong place.

Most of these questions apply to all four products so the FAQ content can largely be shared, with the occasional product-specific addition.

#### 4.8.8 "Learn More" pop-up links

Several helper links on the native configurator (Ultra/Extreme) open pop-up windows (`pop-pages/custpc-ram.htm` etc.). This is a dated UI pattern — modern equivalent is inline expandable helpers, tooltips on hover, or modal overlays that stay in the page context.

Not urgent — the current links work — but worth replacing when next touching the configurator templates. Inline expandable helpers are a much better UX.

#### 4.8.9 What NOT to build

- **Don't try to unify the custom Trader configurator with the native system in a single rebuild.** That's a bigger engineering project than the return justifies. Converge opportunistically over time — when next touching the Trader configs anyway, migrate to whichever system is winning at that point.
- **Don't add more benefit blocks to the Trader PC pages.** Nine is already too many; the goal is fewer, stronger blocks.
- **Don't replicate the landing-page persuasion content on product pages.** The product page job is closing the sale, not re-selling it.
- **Don't build product-specific Trustpilot review filters.** Spreading reviews across four product pages thins each page's review count. Use the aggregate rating instead, with maybe one or two highlighted product-specific reviews pulled manually.

#### 4.8.10 Summary — scope and order

**Phase 1 (this week, ~2 days of work):**
- Add generic helper text to each config option in native ProductCart system, applied to Ultra/Extreme
- Build product-specific helper text alongside if it's quick; drop back to generic if it slows the rollout
- Trim Trader PC / Trader Pro benefit blocks from nine to 3-5
- Add benefit block section to Ultra/Extreme using ported content from trader pages

**Phase 2 (following week, ~2 days):**
- Add trust signals strip to all four product pages
- Add cross-links between product pages
- Add product-specific FAQ sections (4-6 questions each)
- Review Ultra starting CPU / starting price combination

**Phase 3 (future, optional):**
- Migrate Trader PC configurator to native system when next touching for price/option updates
- Replace pop-up "Learn More" links with inline expandable helpers
- Photography additions when Wave 1 completes

This sequencing closes the consistency gap between the product pages quickly (Phase 1), then rounds out the pages with trust/FAQ/cross-link infrastructure (Phase 2), leaving bigger structural changes for later when the cost/benefit is clearer.

---

## 5. Marketing & Offer Opportunities

### 5.1 The bundle offer is undersold
You give away wifi card + speakers (~£60), free cables (~£15 per screen), free delivery, plus discount up to £100. For a six-screen bundle that's £200+ of value. But this is communicated as a small "Save Money, Get Free Cables & Free Delivery with a Bundle" callout.

**Make this a hero message:** "Buy a bundle, save £200+ vs ordering separately, and get a setup that's tested to work together out of the box."

Add a comparison table on the bundles page showing separate-vs-bundle pricing. Quantify the value explicitly.

### 5.2 Trader vs Trader Pro decision tool
Customers want a clear "if you're X, choose this; if you're Y, choose that" decision tool. You have this content scattered through the long-form page; consolidate it.

A four-question quiz on `/trading-computers/`:
1. Which platform do you primarily use? (MT4 / TradingView / NinjaTrader / TradeStation / Bloomberg / Other)
2. How many screens do you want to run? (1-2 / 3-4 / 5-6 / 7+)
3. Do you run heavy multitasking / backtesting? (Yes / No / Sometimes)
4. Budget guidance? (Under £1,500 / £1,500-£2,500 / £2,500+)

→ Recommends a config and captures email address before showing the result. Doubles as a lead-capture and conversion tool.

### 5.3 Free Buyers Guide is buried
This is a Drip lead magnet — your one chance to capture people who aren't ready to buy today. It should have a prominent inline form on every major page and an exit-intent popup. Drip can power remarketing email sequences which would massively help on a high-consideration purchase like this.

### 5.4 5-year hardware cover and lifetime support are buried
Most US competitors offer 3 years; you offer more. Lead with it. These deserve to be top-line offers in hero areas, not buried in product page benefits sections.

### 5.5 Make TraderSpec data more interactive
Static charts are fine but could become a sortable/filterable comparison ("show me CPUs under £200" or "show me what's best for NinjaTrader"). This would generate more SEO surface area and inbound links if you let people share specific comparisons.

### 5.6 Add a "Talk to Darren" booking option
No obvious way to book a call with you. A Calendly link offering a 15-minute consultation could be a powerful conversion tool, particularly for hesitant buyers spending £2,000+. Pre-qualify the appointments by asking budget and use case in the booking form.

### 5.7 Trader-specific content marketing
Beyond the blog, pieces like:
- "What's the best monitor setup for forex day traders?"
- "TradeStation vs NinjaTrader: hardware requirements compared"
- "How many screens do you actually need for futures trading?"
- "Bloomberg Terminal hardware requirements explained"

These would rank for high-intent searches and feed your funnel.

### 5.8 Competitor comparison pages
A "Multiple Monitors vs Scan 3XS" or "vs Falcon" style page is uncomfortable to write but is SEO gold — people search "Scan 3XS vs..." and you want to be the answer. Be factual and fair; the data will speak for itself.

---

## 6. Prioritised Action Plan

### Phase 1 — Quick credibility fixes (1-2 days)
- [ ] Make footer copyright year dynamic
- [ ] Add Trustpilot widget to every page (header strip + footer + product page near CTA)
- [ ] Reconcile prices across all pages, ideally pulling from a single source
- [ ] Upgrade BBC mention to a visible badge in hero areas
- [ ] Replace `javascript:` jump links with proper anchor links

### Phase 2 — Social proof beef-up (1-2 weeks)
- [ ] Pull live Trustpilot reviews into a richer testimonials section
- [ ] Run a customer competition asking for photos of their setups in use
- [ ] Replace trusted-by logos with case-study snippets ("Hedge fund X uses 6x Trader Pros for...")
- [ ] Take photos of the workshop, build process, you (Darren), and the testing rigs
- [ ] Record a 60-90 second video walkthrough of the Trader PC and Trader Pro

### Phase 3 — Page restructures (2-4 weeks)
- [ ] New homepage hero + journey CTAs + trust strip
- [ ] Add comparison table to `/trading-computers/` (Trader vs Trader Pro)
- [ ] Add FAQ accordion to `/trading-computers/` using content from long-form page
- [ ] Rebuild `/computers/` to match `/trading-computers/` quality and structure
- [ ] Update `/pages/trading-computers/` (prices, screenshots, sticky nav, sticky CTA)

### Phase 4 — Offer & conversion mechanics (4-6 weeks)
- [ ] Bundle savings calculator / comparison table on bundles page
- [ ] "Which PC is right for me?" quiz with email capture
- [ ] Inline buyers-guide capture forms on major pages
- [ ] Exit-intent popup for buyers guide
- [ ] "Talk to Darren" Calendly booking option
- [ ] Drip email nurture sequence for buyers-guide subscribers

### Phase 5 — Content & SEO expansion (ongoing)
- [ ] Trader-specific blog content (platform-by-platform hardware guides)
- [ ] Competitor comparison pages
- [ ] Interactive TraderSpec data tools

---

## 7. Things I Did NOT Audit

To set expectations on scope:
- Mobile experience (recommend testing — classic ASP/ProductCart sites can have mobile issues)
- Site speed / Core Web Vitals (worth running through PageSpeed Insights)
- Checkout flow and abandonment
- Google Ads landing page quality scores
- Email capture / nurture sequence content quality
- Blog post quality individually
- Product pages other than the Trader PC
- Monitor / Stand product pages
- Bundles configuration UX in detail
- Schema markup / structured data
- Backlink profile and SEO authority

Worth tackling these in a separate pass once the front-end conversion fundamentals are in place.

---

## 8. What NOT to Change

A few things to deliberately preserve:
- The honest, customer-first technical voice — don't let any redesign sanitise this away
- The TraderSpec.com data presentation (improve it, don't dilute it)
- The detailed customisation options on product pages
- The "we'll talk you out of upgrades you don't need" positioning
- Darren's personal voice on the long-form page (it's authentic and works)

---

## 9. Path to World-Class — `/trading-computers/` Page

This section assumes the v2 mockup (trading-computers2.html) housekeeping items are resolved (quiz references removed, pricing reconciled, comparison defaults set, TraderSpec charts tightened). Everything below is what separates "strong page" from "world-class page" — the level where a trader would send the link to another trader as "this is actually useful, read it."

### 9.1 Philosophical Shift

Strong pages sell. World-class pages educate and earn. The best commercial pages for considered B2B purchases feel less like marketing and more like a knowledgeable friend explaining the decision. The existing `/pages/trading-computers/` long-form page gets close to this in places (the "build it yourself for £25/hour" section, the honest pricing conversation); the new page is more polished but currently more sales-coded. The world-class version finds the balance: structured and scannable like the new page, carrying the knowledgeable-friend voice throughout.

**Test to hold yourself to:** *Would a trader send this page to another trader as "this is actually useful, read it"?* "This is a good sales page" isn't the bar.

### 9.2 Wave 1 — The Authenticity Layer

This is the highest-impact wave. A page with real photography, a real owner on camera, and real customer stories reads as world-class to most visitors even if waves 2 and 3 aren't done yet. The reverse isn't true.

#### 9.2.1 Real photography (see section 3.6 and the imagery discussion)

- Hero photo: actual six-screen array on a real Synergy stand, running real trading platforms (composite approach — real hardware photo + real platform screenshots composited onto screens). Dimmer room, screens as dominant light source, cables visible but tidy. Should look like a real trader's desk first thing in the morning.
- Workshop photos: Darren at the bench. A machine mid-build. The stress-test rig with cables and screens hooked up. A stack of boxes ready to ship.
- Customer setup photos, with permission: range of scales from home trader to prop firm to City office.
- **Captions matter.** Specific captions ("Six-screen Trader Pro running NinjaTrader 8 strategy analyzer — shipped to a futures trader in Guildford, March 2026") are worth more than any number of rendered hero images.

#### 9.2.2 Darren on the page

Place a small professional photo of Darren next to the "Book a 15-min call" CTA. The named-owner CTA is already strong; pairing it with a real face makes it significantly stronger. One of the most under-used trust signals on small specialist sites, and it costs nothing.

#### 9.2.3 Video — specifically, Darren on camera

Conspicuously absent from the mockups. World-class version has at least one, probably two or three short videos embedded.

**Priority video: two-minute "how we spec a trading computer" walkthrough.** Darren-narrated. Shot in the workshop. Picking up components, explaining what matters and what doesn't — CPU for single-thread, GPU just to drive monitors, RAM sized for platform count. Authenticity beats polish: a shaky phone video of Darren in the workshop beats a slick corporate video for this audience.

**Second video: 60-second stress test demo.** Machine on the bench running the 32-hour stress test, thermal readings visible, optional time-lapse. Concrete proof of process.

**Third video (aspirational): customer interview.** Trading firm or individual trader talking about their setup. Harder to organise; not needed on day one but transformative when available.

Video dramatically increases time-on-page (a ranking signal), puts face/voice to the business, and proves claims in ways words can't.

#### 9.2.4 Real social proof throughout

- Replace all placeholder customer logos (Oakwood, Meridian, Strathclyde, etc.) with real logos — with permission — or remove the strip until real ones are gathered.
- Replace styled testimonial cards with the live Trustpilot widget (Business plan required). Real, verifiable, legitimate.
- Dated benchmark data ("Updated March 2026") signals active maintenance — most competitors have 2019-era data still showing.

### 9.3 Wave 2 — The Content Depth Layer

The "knowledgeable friend" layer — stuff only someone with 17 years of customer conversations could produce.

#### 9.3.1 Expanded FAQ pulled from real customer questions

The current FAQ is good but anticipates obvious questions. World-class version anticipates the questions the trader is thinking but won't ask. Examples worth adding:

- *"What happens if my machine dies at 3am before an NFP release?"* — concrete scenario beats generic "lifetime support."
- *"I've been trading for 6 months and I'm not profitable yet. Can I justify £1,500?"* — the honest answer earns enormous trust.
- *"My partner thinks this is overkill — how do I explain it?"* — acknowledging the real conversation traders have.
- *"I'm moving from a Mac. Will I hate Windows?"* — practical, specific.
- *"What's the resale value if I upgrade in 3 years?"* — signals long-term thinking.
- *"Have you ever had a customer you couldn't help?"* — the answer ("yes, and here's what we did") is the ultimate credibility signal.

**Action:** Go through email inbox for the last year and pull out actual questions customers asked. Some will surprise you. The world-class page has answers that feel like they come from 17 years of conversations, not from a marketing brief.

#### 9.3.2 Named competitor comparison

The current benchmark charts reference "Scan 3XS equiv." cautiously. World-class version has a full named comparison section:

"How we stack up against the obvious alternatives" — direct, fair table covering Scan 3XS, CCL, Falcon, Orbital, Blue Aura. Columns:

- Price (for broadly equivalent spec)
- Warranty length
- UK vs overseas
- Trading specialisation (yes/no)
- Published benchmarks (yes/no — only you)
- Phone support
- Typical lead time
- Bundle availability
- Return / money-back guarantee

**Rigorous honesty is the requirement.** If Scan's equivalent is £150 cheaper, say so, and explain what the £150 buys. If Falcon has a more established US reputation, acknowledge it and explain why it doesn't help a UK buyer who needs local support. The table that shows *you losing on one or two dimensions* is the table the reader trusts.

This is the piece most businesses are too nervous to build. Doing it well is a strong confidence signal.

#### 9.3.3 Original research assets

Compound the TraderSpec advantage — this is the one thing competitors genuinely can't copy quickly.

- **Annual "State of Trading Hardware" report.** Everything seen over the year: which platforms got heavier, which CPU generations mattered, GPU requirement changes, common customer-call misconceptions. Downloadable PDF + web page. Gets linked from trading forums and industry newsletters for a year. Weekend of effort.
- **Platform-specific benchmark deep-dives.** "We tested 8 CPUs against NinjaTrader 8 strategy analyzer — here's what actually matters." One per platform, published on TraderSpec, summarised on blog, referenced from `/trading-computers/`. Each is SEO gold for "best CPU for [platform]" type queries.
- **Cost-of-lag calculator.** Interactive widget: trader enters trade frequency and average position size, widget estimates annual cost of 200ms of chart lag vs 50ms. Directionally sensible rather than rigorously accurate. Turns abstract "you need a fast machine" into concrete numbers. Bookmark and share potential.

Any one of these is a world-class-level content asset. All three together puts Multiple Monitors so far ahead of UK competitors that the category wouldn't be comparable.

### 9.4 Wave 3 — The Craft and Speed Layer

Small stuff that makes the page feel considered rather than assembled.

#### 9.4.1 Micro-copy pass

World-class pages sweat the small text. Every body paragraph, button label, tooltip reinforces voice or leaks generic-marketing-speak.

- **Buttons.** "Configure Trader PC" → "Spec my Trader PC." "Book a 15-min call" → "Talk to Darren for 15 minutes." Uses the name already on the page, warmer.
- **Comparison tooltips.** Hovering a 3-star rating: not "Runs fine for moderate use" (generic) but "Fine for one overnight backtest a week — gets slower if you're stacking multiple strategy analyzers" (specific, useful, in-voice).
- **Edge states.** Unsupported configuration in the comparison tool: don't leave cells blank. "This RAM spec isn't offered on the Trader PC — [upgrade to Trader Pro for this option](#)" turns dead-end into conversion path.
- **Section headlines.** Some are strong ("The numbers, because specs without data are marketing.") and some are generic ("Two machines, one decision."). A pass with an editor's eye.

#### 9.4.2 Evidence of ongoing update, not just launch polish

World-class pages feel maintained, not launched.

- **Dates on benchmark data.** "Benchmark data updated March 2026."
- **"Recently asked" module.** Small rotating section showing latest three customer questions with answers. Pulls from email inbox. Feels alive.
- **"Last built" counter.** "47 Trader Pros built in the last 30 days." Social proof + maintenance signal. Only if the number is genuinely impressive — otherwise skip.
- **Visible changelog for the PCs.** "Trader PC 2026.2 — added Core Ultra 7 265KF option, dropped 12th gen i5." Traders love this detail. Almost no one in the category does it.

#### 9.4.3 Page performance

A page selling performance computers *should feel fast.* Run final version through PageSpeed Insights. Targets:

- LCP under 2 seconds
- Total page weight under 1.5MB
- No render-blocking scripts
- Comparison tool responds to dropdown changes in under 100ms

A trading computer page that takes 4 seconds to load and has a janky comparison tool is self-undermining. The reader won't consciously notice but will absolutely feel it. Ironically, this is where many premium-priced competitors fail.

### 9.5 Things to Deliberately NOT Add

A few "world class" temptations worth resisting.

- **Live chat.** Solo operator can't staff it reliably. Nothing kills trust faster than a chat widget that takes 20 minutes or says "we're offline." The "Book a call with Darren" pattern is better for this scale.
- **Full configurator on `/trading-computers/` itself.** Leave that to the product pages. The comparison tool is enough decision-support for this page.
- **Multiple hero variations / heavy personalisation.** A page that guesses whether you're a day trader or hedge fund and shows different content adds brittleness for marginal gain. One strong page beats four mediocre variants.
- **Fancy animations or scroll-triggered effects.** World-class isn't the same as impressive. For a decision-making audience evaluating a serious purchase, clean and fast beats animated and clever.

### 9.6 Sequencing

**Wave 1 is where the biggest jump happens.** A page with real photography, Darren on camera, and real customer stories reads as world-class to most visitors even if waves 2 and 3 aren't done yet. The reverse isn't true: waves 2 and 3 without wave 1 read as "good marketing content," not "real specialist business."

Suggested order:
1. **Wave 1 — Authenticity (4-8 weeks):** Photography session, Darren on camera (one video minimum), real Trustpilot widget, real customer logos/testimonials, dated benchmark data.
2. **Wave 2 — Content depth (2-3 months):** Expanded FAQ from real customer questions, named competitor comparison, cost-of-lag calculator. Original research assets published on blog/TraderSpec over following 6-12 months.
3. **Wave 3 — Craft and speed (ongoing):** Micro-copy pass, performance optimisation, tooltip/empty-state design, rotating "recently asked" module.
