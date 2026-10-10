GRAYSON COULTER - SWIMMING RECORDS AND ACHIEVEMENTS SITE
Public website: v5 (10 October 2026)
Update toolkit: edition 1

QUICK START / VIEWING THE WEBSITE
---------------------------------
Double-click index.html after extracting the ZIP. The site is fully offline:
index.html, nz-records.html, progression.html, consistency.html and assets/.
The four site pages do NOT read the Excel workbook and do NOT need an
internet connection. Keep the assets folder intact. Public GitHub Pages
can serve the four HTML files and assets directly from the repository root.

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
   g. rewrites the four offline HTML files in the REPOSITORY ROOT,
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

9) GitHub Pages publication: commit ONLY the four updated HTML files,
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
  updating; the automated input maps the CURRENT 18 canonical swimmers.
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
