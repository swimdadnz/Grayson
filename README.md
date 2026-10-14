GRAYSON COULTER - SWIMMING RECORDS AND ACHIEVEMENTS SITE
Public website: v5 (10 October 2026)
Update toolkit: edition 1

QUICK START / VIEWING THE WEBSITE
---------------------------------
Double-click index.html after extracting the ZIP. The site is fully offline:
index.html, nz-records.html, progression.html, consistency.html,
personal-bests.html and assets/.
The five site pages do NOT read the Excel workbook and do NOT need an
internet connection. Keep the assets folder intact. Public GitHub Pages
can serve the five HTML files and assets directly from the repository root.

MEDIA GALLERY (INCLUDED IN THIS PATCH)
--------------------------------------
The index page shows ten photographs and one playable MP4 race video.
The following are newly added to assets/:
  Grayson-Coulter-by-Nardia-Mulkerrins-2-1082x720.jpeg
      First sub-4 min 400m Free - QLD Champs 2025
  Grayson-Coulter-by-Nardia-Mulkerrins-3-1082x720.jpeg
      Freestyle breakout - QLD Champs 2025
  2026-10-01 NZSC Scott Rice 200m Free commentary.mp4
      NZSC 2026 - 200m freestyle race commentary
These exact filenames are referenced by index.html and by
update/scripts/build_portable.py when regenerating the site. Keep them in
the root assets/ folder, without renaming. The video is H.264/AAC MP4;
selecting its thumbnail opens a player and attempts to start playback.
On some mobile browsers, pressing the native play control may still be
necessary. For publication, upload the MP4 and JPEGs as normal GitHub files.

UPDATE TOOLKIT - WHERE IT LIVES
-------------------------------
Everything needed for LOCAL updates is packaged in the update/ subfolder:
  update/Grayson_Swim_Site_Master.xlsx      PRIVATE source workbook
  update/input/profiles.csv                 19 SwimCloud source profiles
  update/input/grayson_profiles.csv         Grayson's two profiles
  update/input/Progression_comparison_swimmers.xlsx  comparison roster/DOBs
  update/input/NZ_age_group_margins_review.xlsx       prior benchmark input
  update/incoming_records/current_record_reference.csv  420 original keys
  update/incoming_records/record_changes.csv          REVIEWED new records
  update/scripts/swimcloud_update.ps1       eight chart events across profiles
  update/scripts/swimcloud_grayson_all_events.ps1  all Grayson events
  update/scripts/import_record_changes.py   reviewed NZ record changes/rollover
  update/scripts/update_master.py           import swim history and recalculate
  update/scripts/build_portable.py          rebuild standalone offline site
  update/scripts/build_personal_bests.py     recompute PBs/World Aquatics comparisons
  update/input/world_aquatics_standards.csv  published championship A/B standards (SCM and LCM)
  update/source_templates/                  HTML source layout for rebuilding
  update/UPDATE_SITE.bat                     full Windows workflow
  update/REBUILD_NO_ACQUISITION.bat          records/analysis only
  update/RUN_ACQUISITION_ONLY.bat            browser acquisition, no Python
  update/requirements.txt                   Python library list

PRIVATE DATA / PUBLIC GITHUB - PLEASE READ
------------------------------------------
The master workbook and comparison roster contain exact dates of birth,
including dates of birth of minor swimmers. If this repository is PUBLIC,
committing update/Grayson_Swim_Site_Master.xlsx to GitHub would publish those
details and make them available to anyone, including via Git history.

The supplied update/.gitignore therefore keeps the workbook, raw acquisition
archives, and sensitive spreadsheets in your LOCAL repository checkout but
excludes them from a normal Git commit. This means all update files are in
one local repository folder, as requested, without revealing private data
on the public GitHub Pages site. If you need cloud backup/versioning of the
private files, use a separate PRIVATE GitHub repo or encrypted local backup.
Do not force-add those sensitive files to a public repo.

RECORD-LEADERBOARD PRESENTATION (THIS PATCH)
-------------------------------------------
The NZ Records page displays swimmers holding 5+ current NZ age-group
records by default. Each gender's remaining swimmers are available using
the Show-more disclosure directly below its leaderboard. All records are
retained; the disclosure does NOT exclude lower-ranked swimmers from data.
Estimated current age is calculated using the age group and record dates
at the time of site generation, and may be approximate (especially if
record sources use different age conventions). No date of birth is shared
in the published HTML. Athletes beyond age 18 cannot add age-group records;
their record count may be held or reduced, although rank can improve if
other swimmers lose more records.
The canonical name for 'Samuel Brown' is 'Sam Brown'. This is enforced
by the workbook and update/scripts/name_aliases.py during future imports
and HTML regeneration. All three current records are credited to Sam.
The public Consistency page shows only historical Top-10 counts and rates;
older Top-5 calculation columns remain available in private Excel sheets
for audit purposes and are not included on the public website.

HOW TO REFRESH THE SITE IN THE FUTURE
-------------------------------------
1) Obtain the latest Swimming NZ age-group SCM and LCM record books.
   Save the PDFs in update/incoming_records/ as reference documents.
   Covers individual male/female age-group records 13-18, not Open or relays.
   IMPORTANT: the scripts do NOT automatically parse/interpret PDF tables.

2) Compare the new record books to
   update/incoming_records/current_record_reference.csv and the workbook's
   05_Record_Benchmarks sheet. Enter ONLY CHANGED RECORDS in
   update/incoming_records/record_changes.csv. Leave all other rows absent.
   Input columns (text, not Excel date/time values):
       Record Key        e.g. NZ|SCM|M|14|200 Free
       New Holder        e.g. First Last (for ties: First Last / Another Name)
       New Time          e.g. 1:48.13
       New Date          YYYY-MM-DD, the date of the new record swim
       Source URL        official result / record-book link
       Next Fastest Name optional manual override
       Next Fastest Time optional manual override, e.g. 1:50.40
       Review Notes      optional
   IMPORTANT: 'New Time' must be plain text (1:48.13), not an Excel
   elapsed-time or clock-time cell. Use the CSV in a plain text editor or
   explicitly format Excel columns as TEXT before entry/export.

3) Record change import rules:
   - A new DIFFERENT holder who swims faster displaces the former holder:
     the former holder/time becomes the current next-fastest candidate.
   - The SAME holder improving their own time retains the existing
     next-fastest DIFFERENT swimmer; the old PB is NOT the comparator.
   - Tied record holders are each credited in the leaderboard but a tied
     record gives a zero marginal advantage versus another holder.
   - Current record history is retained and the old current entry closed.
   - Ambiguous former ties require manual verification; no silent guessing.
   - A reviewed explicit next-fastest name AND time overrides an inferred
     comparator (enter both or neither).
   - 'Standard' benchmarks are not swimmer-held age-group records.
   Review any new Top-20 positions manually against the next-fastest
   all-time list; this step is NOT fully automatable from the current data.

4) Update the master comparison roster if required. If a swimmer acquires
   an additional SwimCloud profile, add the mapping carefully to the
   swimmer source dimension and acquisition profiles.csv. Grayson's two
   known SwimCloud IDs already map to the same canonical person. DOB and
   source IDs must be checked before acquiring new swimmers.

5) On Windows, double-click update/UPDATE_SITE.bat. It:
   a. backs up the private workbook in update/archive/;
   b. applies changes entered in record_changes.csv (archiving that file);
   c. opens Chrome/Edge to acquire all eight chart events for the
      comparison population; you clear Cloudflare and PRESS ENTER;
   d. imports fresh times without discarding historical races;
   e. opens Chrome/Edge for BOTH Grayson profiles; you again clear
      Cloudflare and PRESS ENTER to retrieve all listed events;
   f. recalculates the workbook and generates fresh internal JSON;
   g. rewrites the five offline HTML files in the REPOSITORY ROOT,
      including record leaderboards, next-fastest margins, progression,
      consistency, heat maps and new page creation date.
   The browser acquisition installs no Chrome extensions, Playwright or
   Python packages. It may pause for additional browser confirmations or
   fail if SwimCloud changes access controls or rate-limits requests.

6) Requirements for the workbook/chart REBUILD step (not acquisition):
   Python 3 plus openpyxl, numpy, scipy and beautifulsoup4. One-time install,
   if needed: py -m pip install -r update/requirements.txt
   If Python cannot be installed now, run update/RUN_ACQUISITION_ONLY.bat
   to acquire the two ZIP files without updating the public pages. Preserve
   those output ZIP files; they can later be imported using Python or sent
   back to ChatGPT for rebuilding.

7) If only the NZ record input or manual workbook data has changed, run
   update/REBUILD_NO_ACQUISITION.bat. It applies any staged record changes
   and recalculates the portable HTML, WITHOUT visiting SwimCloud.
   If an acquisition fails, do not commit partially refreshed pages.
   Restore from update/archive/ if needed, then retry.

8) Review before publishing:
   - QA_Checks sheet: referential integrity and any REVIEW flags.
   - 05_Record_Benchmarks: holders, times, dates, verified comparators.
   - 06_Record_History and 07_AllTime_Candidates: correct rollover.
   - A_Outstanding_Records: male and female Top 20, check any new margins.
   - Progression: refreshed stars, historical swimmers stable, Grayson
     combined once; exact age based on PRIVATE DOB.
   - Consistency: all Grayson event/course histories from 1 Jan 2024,
     historical Top 5 / Top 10 at the time of each race; trend fits.
   - Heat maps: anniversaries measured back from HTML generation date;
     one current record/event/age/course each, ties counted only once.
   - Recency narrative is descriptive; re-check attributed swimmer trends
     as record books change rather than accepting old conclusions blindly.
   - Open the site locally (file://) and, ideally, test on a phone as well.
   - Check all ten photographs and the embedded MP4 player.

9) GitHub Pages publication: commit ONLY the five updated HTML files,
   assets/ and non-sensitive update scripts/README (optionally the public
   source_templates folder). Leave workbook, raw ZIPs, DOB roster, and
   archive files outside the PUBLIC Git history. Upload/push and view the
   Pages URL. You may keep the entire update folder locally under the
   cloned repository working tree, even though private files are gitignored.

IMPORTANT LIMITATIONS / SOURCE DATA
-----------------------------------
- NZ record PDFs require manual extraction and human approval. New PDFs
  placed in incoming_records/ are NOT consumed automatically by the script.
- Updating only SwimCloud cannot itself update NZ records or the all-time
  next-fastest benchmarks; those require official source comparison.
- SwimCloud may not show a new meet or every heat/final immediately; review
  especially the newest NZSC meet against official published results.
- Historical times are deduplicated by the underlying workbook importer;
  ambiguous equal-time, same-date races should be manually checked.
- Age-group record analysis is ages 13-18, SCM and LCM, as in the original.
- The six trendline candidate models and residual statistics are descriptive,
  not predictions of eventual best performances.
- After adding NEW swimmers to the roster, inspect the importer before
  updating; the automated input maps the CURRENT 19 canonical swimmers.
- This update toolkit was tested for offline rebuild, but the Windows
  acquisition steps require a live authenticated browser and SwimCloud
  availability; those cannot be validated in this packaging environment.

DID THE PUBLISHED PAGES CHANGE IN THIS DELIVERY?
-----------------------------------------------
For this gallery-only patch, index.html, README.txt and
update/scripts/build_portable.py have changed. Two JPEGs and one MP4 are new.
The NZ Records, Progression and Consistency pages, existing eight photos,
Excel master database and acquisition scripts have NOT changed. You can
apply only the files in this patch to an existing v5 GitHub repository.

PERSONAL BESTS PAGE - ADDED OCTOBER 2026
--------------------------------------
Public file: personal-bests.html. The Home navigation links to it without
changing the three feature cards on the landing page. The table covers:
- LCM: 50/100/200/400 Free, 50/100 Fly, 200 IM (7 events).
- SCM: 50/100/200/400 Free, 50/100 Fly, 100/200 IM (8 events).
It uses the fastest legal performance for Grayson from 04_Swims, including
valid relay lead-offs, across BOTH SwimCloud profiles. Equal PBs display
the first recorded date and meet. Times are course-specific and never
converted between SCM and LCM. If a new meet appears in the SwimCloud
acquisition, the fastest qualifying time is selected when the page rebuilds.

Published World Aquatics benchmarks (men's individual events):
- LCM: World Aquatics Championships Budapest 2027, 50m qualifying A/B.
  https://mediamanager.ws/media/pages/s/swimmsvk.sk/2026/06/sw-qualification-world-aquatics-championships-budapest-2027.pdf
- SCM: World Aquatics Swimming Championships Beijing 2026 (25m) A/B.
  https://resources.fina.org/fina/document/2026/02/16/2bee1e7b-bb50-4d97-855a-650661073eb0/Qualification-World-Aquatics-Swiming-Championships-25m-Beijing-2026.pdf
A/B columns use standards for the SAME course as the swimmer's PB.
Gap = (PB seconds / published standard seconds - 1) * 100; positive is
slower than the benchmark, negative is faster. Championship standards
are NOT age-specific, and achievement of a standard is not automatic
selection or official qualification. Approved meet and qualifying period
requirements apply.

For each future refresh:
1. Run the existing Windows acquisition and update process; the
   `build_portable.py` step calls `build_personal_bests.py` automatically.
   No extra interactive SwimCloud acquisition is needed for this page.
2. Confirm `A_Personal_Bests` and `World_Aquatics_Standards` appear in the
   PRIVATE master workbook and compare any recent PB dates/meets with
   the imported SwimCloud records. The new page is always built offline.
3. If World Aquatics publishes new A/B standards for either course,
   manually update `update/input/world_aquatics_standards.csv`, including
   the championship name, source URL, and qualifying period. The next
   rebuild updates both the workbook and page. Do not silently reuse
   obsolete entry standards for a different championship.
4. Publish `personal-bests.html` along with `index.html` whenever their
   output changes. Do not publish the private workbook or any source CSV
   containing private swimmer information.

The PB spreadsheet audit columns use Excel formulas pointing to source
rows in 04_Swims, plus a separate World_Aquatics_Standards input sheet.
Excel recalculates formula outputs on open. The public HTML is computed
directly from numeric race times so it works under file:// without Excel.


PROGRESSION: AGE-YEAR SEASON BESTS
--------------------------------
The Progression page compares age-year bests for Grayson and the selected
swimmer in the event selected above the chart (ages 11 through 20).
This is NOT a calendar-year or standard swimming-season definition: an
age-year runs from one birthday until the next. Exact age is calculated
privately from DOB and swim date and is already present in the generated
progression JSON as the decimal age. The lowest valid time from each age
interval is displayed. Missing coverage appears as "-", not as zero.
Percentage difference = (Grayson / selected - 1) * 100. Positive means
Grayson is slower; negative means faster. Comparing Grayson with himself
gives 0.00% where data exist.

The table and chart use the SAME acquisition and swimmer profile mapping.
No new master workbook column or extra script is needed. Every rebuild
reads update/source_templates/progression.html and replaces the embedded
progressionData with the latest generated_data/progression.json; the
age-year table updates from those records automatically.

FIVE-PAGE NAVIGATION AND PERSONAL BESTS TEMPLATE
-----------------------------------------------
NZ Records, Progression and Consistency now link to Personal bests in the
header, matching the existing links on Home and Personal bests. The
source_templates copies are updated so the links survive site refreshes.
The Personal bests page now also has its own source template at:
  update/source_templates/personal-bests.html
The build_personal_bests.py script reads that template, preserving its
layout and navigation while replacing the LCM/SCM tables, standards
methodology, refresh dates and footer on each rebuild. Do not edit the
public personal-bests.html alone if you intend the change to persist.

This changes-only release also includes the previously uncommitted
Progression season-bests table and its matching source template.

OCTOBER 2026: HISTORICAL AUSTRALIAN COMPARISONS AND OTHER RECORD CATEGORIES
--------------------------------------------------------------------------
The progression database now includes three Australian historical swimmers
with a one-off data reconciliation against Swimming Australia's boys' age
all-time Top 10 lists, dated 28 August 2026 (LCM) and 24 September 2026 (SCM).
These lists are age-specific Top 10 entries, NOT complete season histories.

Source CSV: update/input/australian_historical_top10_verified.csv
One-off importer: update/scripts/import_historical_top10_once.py
Use the importer ONLY to apply corrections to the historical CSV; it matches
swimmer + date + event/course + time and will not reinsert exact duplicates.
65 published entries were evaluated: 2 matched already-acquired SwimCloud swims,
63 were added. The workbook retains Swimming Australia as source for additions.
Ian Thorpe has been added with SwimCloud profile https://www.swimcloud.com/swimmer/939756/times/
and his correct date of birth in the PRIVATE workbook. This does not mean his
entire SwimCloud racing history has been acquired or independently verified.
Ages are computed privately from DOB and official swim dates. This does not
change Grayson's own data or consistency calculations.

On Progression, an age dropdown (11-20) now controls a ranked age-year table for
the selected event/course. One fastest recorded swim per swimmer is ranked,
including Grayson when he has an age-eligible result. Swimmers without times
at that age are omitted from the ranking. The separate two-swimmer season-best
comparison below the chart is retained. Missing data should never be treated
as evidence that a swimmer did not swim that event.

Historical champion data is considered fixed: the selected-events SwimCloud
PowerShell acquisition now refreshes only current age-group stars. Grayson's
two profiles continue to be refreshed separately by the all-events acquisition.
The historical Australian Top-10 supplements are preserved in the master when
the normal UPDATE_SITE.bat is run. Do not reacquire historical champions by
routine. If a historical correction is needed, review the official source,
change the historical CSV, run the one-time importer, then rebuild.

Three additional categories are now displayed under NZ age-group record totals:
 - Queensland All-Comers: 10
 - Australian All-Comers: 3
 - Unofficial international age bests: 3
Each has an expandable table. Some swims qualify in multiple categories.
The international age bests refer to Dominique Eisold's private collection,
not World Aquatics-ratified records.

Manual refresh process for QLD/AUS/world record categories:
1. Check the newest official Queensland All-Comers and Australian All-Comers
   record listings, and the international age-records website. Determine
   whether Grayson still holds each mark and whether any new category record
   has been set. Do not assume an old result remains current indefinitely.
2. Edit update/input/other_record_categories.csv (one row per record-category
   combination). Preserve the Category values QLD, AUS, or WORLD; enter the
   event, course, age, official result time, swim date, source URL and date
   verified. Do not confuse an Australian AGE record with an All-Comers record.
3. Run update/REBUILD_NO_ACQUISITION.bat (or the full UPDATE_SITE.bat if new
   SwimCloud results are also needed). build_portable.py reads the category
   CSV directly; do not hard-code the homepage card numbers.
4. Verify the public index.html totals and expanded lists against official
   source documents, the NZ tables, and the private master before publishing.
5. Keep update/Grayson_Swim_Site_Master.xlsx and input/profiles.csv (which
   both contain date-of-birth data) out of PUBLIC GitHub history. Keep the
   private update folder in a PRIVATE repository or an ignored local directory.

Data limitations: published Australian historical lists are only Top-10 ages
13-18; only the requested eight event/course combinations were included.
Actual age-year rankings may be incomplete for historical champions. Ian
Thorpe's full profile cannot currently be automatically scraped from this
build environment. All calculations are descriptive, not predictive.

SCM RATIFICATION + HISTORICAL QUEENSLAND AGE-12 SWIMS (15 OCT 2026)
-----------------------------------------------------------------
The supplied, RATIFIED Swimming New Zealand Short Course Record Book is dated
6 October 2026. All 216 individual age-group records (ages 13-18, both genders)
in that book have been reconciled against the master workbook: 216/216 record
times and current holders match after normalising existing aliases such as
Samuel Brown -> Sam Brown and Te Paki -> Te Rauroha Paki. No change to current
NZ record counts, next-fastest comparators or record times was required.
The master workbook now identifies the 6 October 2026 Swimming NZ list as
the ratified reference, rather than relying on the earlier meet editorial
report to confirm current records. The official reference PDF is retained in
  update/incoming_records/Swimming_NZ_Short_Course_Records_as_at_2026-10-06.pdf
and the 216-row audit transcription is in
  update/input/ratified_scm_age_records_2026-10-06.csv
Official current-record portal: https://www.swimmingnz.org/performance-and-coaching

Five previously missing age-12 LCM performances from the 2024 Queensland
Championships have been added to Grayson's PRIVATE 04_Swims dataset:
  15 Dec 2024, 200 Free  1:54.40
  16 Dec 2024, 100 Fly     58.52
  16 Dec 2024, 200 IM   2:11.16
  17 Dec 2024, 50 Free    24.52
  18 Dec 2024, 100 Free   53.25
These are sourced from Swimming Queensland's All-Comers record list and have
NOT been invented as SwimCloud races. The record publication does not
establish each heat/final round, so Round is recorded as TIM (unspecified).
The source file is update/input/queensland_age12_historical_2024.csv.

The two included data changes have been APPLIED ALREADY. The one-off command
below is supplied for audit, recovery or use after restoring an older workbook:
  cd update
  python scripts/reconcile_2026_scm_and_queensland_once.py
It is intentionally NOT part of UPDATE_SITE.bat or REBUILD_NO_ACQUISITION.bat.
It checks the full official age-group record book before changing any record
provenance, matches 2024 Queensland swims on swimmer/course/event/date/time,
and safely skips performances already in the database. Should a matching
SwimCloud race later appear with a different meet label or round, the normal
acquisition importer also avoids duplicating these five supplements.

Because the additions are historical, they do not change current NZ records
or Grayson's current PBs. They DO improve his age-12 LCM progression points,
the age-year season best and ranked comparisons, and his 2024-onwards
consistency statistics. The current publication's existing 'as at' date is
preserved in this narrow data-reconciliation patch to avoid unrelated changes;
future regular refreshes will set a new date on all pages automatically.

For ALL FUTURE updates, continue obtaining the latest Swimming NZ SCM/LCM
ratified records and staging genuine changes in record_changes.csv. Review
Queensland All-Comers, Australian All-Comers and privately compiled world
age bests separately in update/input/other_record_categories.csv. These
historical Queensland times should not be manually appended a second time.
