/**
 * Post-Onboard Tracker — All-in-One Apps Script Web App
 * ======================================================
 * Everything runs inside this one Google Apps Script project:
 *   - The HTML form (Index.html) the user interacts with
 *   - HubSpot reads (contact + deal + salesperson + source)
 *   - Fathom reads (onboard date + onboarder)
 *   - The write into the tracking sheet
 *   - A debug helper to list a contact's existing tasks (so we can see
 *     real task titles created by sequence enrollment, before we build
 *     the due-date-fixing logic)
 *
 * ── SETUP (one time) ──────────────────────────────────────────────
 * 1. Project Settings → Script Properties → add:
 *      HUBSPOT_TOKEN  =  your HubSpot private app token (pat-na2-...)
 *      FATHOM_API_KEY =  your Fathom API key
 * 2. Deploy → New deployment → Web app
 *      Execute as: Me
 *      Who has access: Anyone with Google account
 * 3. Open the web app URL. Done.
 */

// ── Config ────────────────────────────────────────────────────────
var SHEET_ID = "YOUR_PRODUCTION_TRACKING_SHEET_ID_HERE"; // PRODUCTION sheet — replace with your own Google Sheet ID
var TAB_NAME = "Edit This One!";

var REFERRAL_SOURCES = [
  "Active User Referral", "Sales Referral", "Other-Referral", "Other Referral"
];

var TEAM_NAMES = [
  "Andrew Stein","Collin Michels","Ben Ogan","Sam Absalom",
  "Brock Baker","Tristan Nelko","Peter Billing","Andrew Yip",
  "Pierce Gregory","Chase Roberts","Sam Stinger","Charlie Coppola"
];

// Peter Billing's HubSpot owner ID — he always gets the thank-you card
// task, so this is used directly rather than doing an owner lookup for him.
var PETER_BILLING_OWNER_ID = "YOUR_HUBSPOT_OWNER_ID_HERE"; // replace with the real HubSpot owner ID for whoever always receives this task

// How many days back the Fathom fallback scan looks. This only affects
// contacts whose Fathom invitee filter misses them (e.g. two people sharing
// one booking) — normal contacts resolve instantly regardless of this value.
//   90 = TESTING already-onboarded people whose demo is several weeks old.
//   30 = PRODUCTION (fresh onboards only) — faster, fewer pages to scan. [CURRENT]
var FATHOM_SCAN_DAYS_BACK = 30;

// ── Fathom rate-limit controls ──────────────────────────────────────
// The Fathom fallback scan (see getFathomMeetingRange) fetches each
// candidate rep's meeting history one at a time, SEQUENTIALLY — not
// concurrently. Two earlier versions tried concurrent fetching (first
// firing all reps at once per round, then capped to 3-at-a-time after
// that tripped Fathom's rate limit), but live testing showed even 3
// concurrent requests still got rate-limited on nearly every round —
// Fathom's real enforcement doesn't tolerate multi-request bursts the way
// a simple "60 calls/minute budget" would suggest. Going fully sequential,
// paced FATHOM_PACE_MS_PER_REQUEST apart between each individual request
// (the same pattern fetchFathomPages() already uses elsewhere in this
// file, which has not shown this problem), is what actually stays inside
// Fathom's real limit.
var FATHOM_PACE_MS_PER_REQUEST = 1100; // matches the ~1100ms/request pacing used elsewhere in this file for a safe rate

// How long a contact's Fathom/HubSpot-meetings resolution stays cached.
// Re-pulling the SAME contact within this window skips the slow Fathom scan
// entirely and reuses the last result — this only speeds up REPEAT pulls
// (e.g. re-testing the same contact), not a contact's first-ever pull. Kept
// short since a genuinely new Fathom recording should be picked up quickly
// once you re-test after it's made.
var PULL_CACHE_TTL_SECONDS = 300; // 5 minutes

// Bumped whenever a real Fathom-resolution-logic change ships. Included in
// the cache key below so old cached results automatically become
// irrelevant the moment the code actually changes — no manual "bypass
// cache" step needed. Confirmed real problem this solves: testing a fresh
// fix by re-pulling the same contact within 5 minutes used to silently
// replay the OLD (pre-fix) result, making a working fix look like it had
// no effect. Bump this string any time resolveFathomAndFinalize,
// getFathomMeetingRange, or getMeetingRepsFromHubSpot's logic changes.
var CODE_VERSION = "2026-08-19-prod1";

// Master on/off switch for the AI accuracy check. This is the only feature
// in Lupe that costs money (real Anthropic API calls) — everything else
// (HubSpot, Fathom, the sheet, caching, the trace) is free regardless of
// this setting. Set to false to run Lupe entirely free; the trace/AI panel
// will just say the check is disabled instead of attempting any API call.
// Flip back to true once there's Anthropic credit on the account.
var AI_ACCURACY_CHECK_ENABLED = false;

// Cache is scoped to the Fathom scan result ONLY — never the sheet-existence
// check or AppData lookup, both of which must always reflect the current
// state (e.g. right after an insert) and are cheap Sheet reads anyway, not
// the bottleneck. HubSpot Meetings-tab data isn't cached here either — it's
// fetched once in the fast path and reused, never re-fetched.
function getCachedFathomResolution(key) {
  try {
    var raw = CacheService.getScriptCache().get(key);
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    return null; // cache errors are non-fatal — just recompute fresh
  }
}
function setCachedFathomResolution(key, bundle) {
  try {
    CacheService.getScriptCache().put(key, JSON.stringify(bundle), PULL_CACHE_TTL_SECONDS);
  } catch (e) {
    // non-fatal — this pull's result just won't be cached for next time
  }
}

// The ONLY people who may appear as a salesperson or onboarder. Any resolved
// name not on this list is rejected and the field is left blank for manual
// entry — this stops cold-callers/bookers (e.g. "Cold Call - Andrew Yip")
// from ever leaking in from the deal Source field or a mis-hosted meeting.
// Edit this list to add/remove valid reps (use their exact HubSpot name).
var VALID_REPS = [
  "Ben Ogan", "Sam Stinger", "Sam Absalom", "Collin Michels",
  "Brock Baker", "Charlie Coppola", "Andrew Stein"
];

// Exposes the roster to Index.html so the Onboarder/Salesperson dropdowns
// pull from this SAME array — the single source of truth already used
// everywhere else (validRep, UNAMBIGUOUS_REP_FIRST_NAMES, etc.) — instead
// of a second, hardcoded copy in the frontend that could silently drift
// out of sync if the roster ever changes here without a matching edit
// there.
function getValidRepsList() {
  return VALID_REPS;
}

// Email for each valid rep — used to narrow the Fathom fallback scan
// server-side via recorded_by[], instead of paging through the whole
// workspace's meeting history. Any call NOT recorded by one of these people
// gets rejected by validRep()/resolveHost() downstream anyway, so this
// filter can never lose a result that would have counted — it only cuts
// out meetings we'd have discarded regardless (a cold-caller's bookings,
// an unrelated staff member's calls), making the scan itself faster.
var REP_EMAILS = {
  "Ben Ogan": "ben.ogan@creonesource.com",
  "Andrew Stein": "andrew.stein@creonesource.com",
  "Charlie Coppola": "charlie.coppola@creonesource.com",
  "Brock Baker": "brock.baker@creonesource.com",
  "Sam Stinger": "sam.stinger@creonesource.com",
  "Sam Absalom": "sam.absalom@creonesource.com",
  "Collin Michels": "collin.michels@creonesource.com"
};

// Returns the canonical roster name if `name` matches a valid rep, else "".
function validRep(name) {
  if (!name) return "";
  var n = String(name).trim().toLowerCase();
  for (var i = 0; i < VALID_REPS.length; i++) {
    if (VALID_REPS[i].toLowerCase() === n) return VALID_REPS[i];
    if (n.indexOf(VALID_REPS[i].toLowerCase()) >= 0) return VALID_REPS[i]; // handles "Sam Absalom <email>" etc.
  }
  return "";
}

// Unambiguous rep first names, computed from VALID_REPS: a first name is
// only usable for name-only title matching (e.g. a call titled "...and
// Andrew..." with no last name) when EXACTLY ONE roster member owns that
// first name. "Sam" is shared by Sam Stinger AND Sam Absalom, so it's
// deliberately excluded — matching on "Sam" alone can't tell which of the
// two is meant, whereas "Ben", "Collin", "Brock", "Charlie", "Andrew" each
// belong to only one person and are safe to match on their own. Computed
// (rather than hardcoded) so adding/removing reps in VALID_REPS keeps this
// correct automatically.
function computeUnambiguousRepFirstNames() {
  var counts = {};
  VALID_REPS.forEach(function(full) {
    var first = full.trim().split(/\s+/)[0].toLowerCase();
    counts[first] = (counts[first] || 0) + 1;
  });
  var map = {}; // firstName(lc) -> canonical full name
  VALID_REPS.forEach(function(full) {
    var first = full.trim().split(/\s+/)[0].toLowerCase();
    if (counts[first] === 1) map[first] = full;
  });
  return map;
}
var UNAMBIGUOUS_REP_FIRST_NAMES = computeUnambiguousRepFirstNames();

// ── Accuracy trace ──────────────────────────────────────────────────
// Collects a step-by-step record of what pullContactData actually checked —
// which sources it read, what it found, and whether each candidate passed
// roster validation — so the frontend can render an audit trail alongside
// the result. Reset at the start of every pullContactData call.
//
// IMPORTANT LIMITATION: google.script.run is a single request/response
// call, not a streaming connection, so this trace cannot update live in
// the UI while the backend is running. It's assembled during the call and
// delivered in full once pullContactData returns.
var TRACE = [];
// Wall-clock start time for the CURRENT trace, reset alongside TRACE at
// the top of pullContactData and resolveFathomAndFinalize. Every trace
// line gets a "[t+Nms]" suffix showing elapsed time since that reset —
// this turns a pasted trace from a list of WHAT happened into one that
// also shows WHEN, in real milliseconds. Appended directly into the
// visible detail string (not a separate JSON field) so it shows up in
// the existing trace UI with no changes needed on the Index.html side.
var TRACE_START_MS = 0;
function traceLog(step, detail, status) {
  // status: "info" (default) | "ok" | "warn" | "fail"
  var detailStr = detail ? String(detail) : "";
  if (TRACE_START_MS) {
    detailStr += " [t+" + (new Date().getTime() - TRACE_START_MS) + "ms]";
  }
  TRACE.push({ step: String(step), detail: detailStr, status: status || "info" });
}

// ── Web app entry point ───────────────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("LUPE")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

// ── Secrets ───────────────────────────────────────────────────────
function getProp(key) {
  var v = PropertiesService.getScriptProperties().getProperty(key);
  if (!v) throw new Error("Missing Script Property: " + key + ". Add it under Project Settings → Script Properties.");
  return v;
}

// ── Helpers ───────────────────────────────────────────────────────
function extractContactId(input) {
  if (!input) return null;
  var s = String(input).trim();
  var patterns = [
    /\/contact\/(\d+)/,
    /\/record\/0-1\/(\d+)/,
    /contactId=(\d+)/,
    /\/contacts\/\d+\/contact\/(\d+)/,
    /\/contacts\/(\d+)(?:[/?]|$)/
  ];
  for (var i = 0; i < patterns.length; i++) {
    var m = s.match(patterns[i]);
    if (m) return m[1];
  }
  if (/^\d+$/.test(s)) return s; // plain ID
  return null;
}

// Explicit timezone — matches CRE OneSource's HubSpot account setting.
// This matters because Fathom returns timestamps in UTC, and using
// getMonth()/getDate()/getFullYear() directly relies on the Apps Script
// project's ambient default timezone, which may not match your actual
// business timezone. That mismatch was silently shifting dates by a day
// for meetings near midnight in either zone. Utilities.formatDate pins
// the conversion to a known, explicit zone so this can't drift.
var ORG_TIMEZONE = "America/Denver";

function fmtDate(d) {
  var dt = (d instanceof Date) ? d : new Date(Number(d));
  if (isNaN(dt.getTime())) return "";
  return Utilities.formatDate(dt, ORG_TIMEZONE, "M/d/yy");
}

function isoDate(d) {
  var dt = (d instanceof Date) ? d : new Date(Number(d));
  if (isNaN(dt.getTime())) return "";
  return Utilities.formatDate(dt, ORG_TIMEZONE, "yyyy-MM-dd");
}

// Earlier/later of two "YYYY-MM-DD" strings — this format compares
// correctly as plain strings (lexicographic order = chronological order),
// no Date object parsing needed. Either argument may be blank/missing.
function earlierIsoDate(a, b) {
  if (!a) return b || "";
  if (!b) return a || "";
  return a < b ? a : b;
}
function laterIsoDate(a, b) {
  if (!a) return b || "";
  if (!b) return a || "";
  return a > b ? a : b;
}

// Reformats an already-correct "YYYY-MM-DD" string into "M/d/yy" using
// plain string splitting — NO Date object round-trip.
//
// Why this exists: a bare date-only string like "2026-06-30" gets parsed
// by `new Date(...)` as UTC MIDNIGHT (per the JS spec), not local midnight.
// Converting that UTC-midnight instant back into America/Denver time (which
// is behind UTC) rolls it back to June 29th. Since the incoming string is
// already the correct calendar date (produced by isoDate() above, which
// DID do the timezone conversion correctly against a real timestamp), all
// we need here is to re-punctuate the string — reparsing it through Date()
// re-introduces the exact bug we already fixed once.
function isoStringToDisplayDate(isoStr) {
  if (!isoStr) return "";
  var parts = String(isoStr).slice(0, 10).split("-");
  if (parts.length !== 3) return "";
  var y = parts[0], m = parseInt(parts[1], 10), d = parseInt(parts[2], 10);
  if (!y || !m || !d) return "";
  return m + "/" + d + "/" + y.slice(2);
}

// ── Duplicate-recording collapse ────────────────────────────────────
// When more than one of our reps joins the same customer call with their
// own Fathom bot recording enabled (e.g. both Ben Ogan and Sam Stinger on
// the same Zoom), Fathom creates a SEPARATE recording per bot for what is
// really ONE real call — same title, same/near-identical start time,
// different recorded_by. Left uncorrected, this double-counts the "# of
// onboarding calls" field. Two meetings are treated as the same real call
// if their titles match (case-insensitive) and their start times are
// within DEDUPE_WINDOW_MIN minutes of each other; only the first one
// encountered is kept, the rest are dropped as duplicate bot recordings.
var DEDUPE_WINDOW_MIN = 60;
function dedupeMeetings(items, getTimestamp) {
  var kept = [];
  items.forEach(function(m) {
    var title = ((m.meeting_title || m.title) || "").trim().toLowerCase();
    var ts = getTimestamp(m);
    var isDupe = false;
    for (var i = 0; i < kept.length; i++) {
      var k = kept[i];
      var kTitle = ((k.meeting_title || k.title) || "").trim().toLowerCase();
      if (title && kTitle === title && Math.abs(getTimestamp(k) - ts) <= DEDUPE_WINDOW_MIN * 60 * 1000) {
        isDupe = true;
        break;
      }
    }
    if (!isDupe) kept.push(m);
  });
  return kept;
}

// ── Later-touchpoint guard ──────────────────────────────────────────
// Titles for post-onboarding SOP calls (90-day renewal, refund check-ins,
// upsell, etc.) that must NEVER be miscounted as an onboarding call, even
// if they happen to mention the contact's name and a rep's name together.
// "next steps" specifically covers a real observed false-positive: a call
// titled "OneSource Next Steps (Andrew & Andrew)" — a pre-decision
// wrap-up after a demo, not an actual hands-on onboarding session — got
// wrongly classified as onboarding via the generic "[Client] and [Rep]"
// fallback, because the CONTACT himself happened to also be named Andrew,
// so his own first name satisfied both the "contact present" and the
// "rep present" checks simultaneously.
// "revisit"/"reconnect" cover the same failure mode from a different real
// case: "OneSource Revisit (Brett & Andrew)" was wrongly classified as
// the earliest training call (via the same generic fallback — both names
// present) when it's actually a PRE-onboarding sales-cycle touchpoint,
// months of drift aside from the real training call. This pushed the
// reported "first onboard date" to the wrong, earlier call.
// NOTE: "best practices" was briefly added here and then REVERTED — it
// was a wrong fix. "OneSource Best Practices" calls are genuine
// onboarding sessions (confirmed directly: Fathom tags them "STE ONBD"
// internally). The real issue for Contact H's case was the OPPOSITE:
// "OneSource Click-Around" was being trusted as onboarding evidence when
// it's actually a pre-sale DEMO for a prospect (Fathom tags it "STE
// DEM"). See the meeting_type handling below, which uses Fathom's own
// tag as the authoritative signal instead of guessing from title text.
var LATER_TOUCHPOINT_KEYWORDS = [
  "renewal", "refund", "check in", "check-in", "checkin",
  "upsell", "expectations", "90 day", "90-day", "favorite part",
  "next steps", "revisit", "reconnect"
];
function isLaterTouchpointTitle(title) {
  for (var i = 0; i < LATER_TOUCHPOINT_KEYWORDS.length; i++) {
    if (title.indexOf(LATER_TOUCHPOINT_KEYWORDS[i]) >= 0) return true;
  }
  return false;
}

// ── Onboarding-call detection ──────────────────────────────────────
// Onboarding calls show up under a lot of different title patterns in
// practice — reps don't title calls consistently. Confirmed real examples:
// "training call", "OneSource Training - ...", "STE ONBD",
// "CRE OneSource Onboard Call with ...", and plain "[Client] and [Rep]" /
// "[Client] and [Rep] Call 2" with NO keyword at all. Demos ("OneSource
// Meeting", "STE DEMO") and later SOP touchpoints (renewal, refund, etc.)
// must NEVER be counted as an onboard.
//
// contactFirstName (optional) enables the generic "[Client] and [Rep]"
// fallback — if the meeting title contains both the contact's own first
// name AND a rep's name, and it isn't a demo or a later touchpoint, it's
// treated as onboarding even with no other keyword. The rep match accepts
// either a FULL roster name (e.g. "Andrew Stein") OR, when the title only
// gives a first name (e.g. "Elia Sanchez and Andrew"), an UNAMBIGUOUS rep
// first name — one that belongs to exactly one roster member. "Sam" is
// intentionally excluded from the first-name-only path since two reps
// (Sam Stinger, Sam Absalom) share it and a bare "Sam" can't tell them
// apart; a title would need to spell out which Sam. Passed only where the
// caller has the contact's resolved name in scope; omitted call sites
// (debug tools) simply skip that fallback and keep the keyword-only check.
function isOnboardingMeeting(m, contactFirstName) {
  // Fathom's own meeting_type tag, when present, is the AUTHORITATIVE
  // signal — the team's own explicit classification of the call,
  // confirmed directly (not inferred): "STE ONBD" is the ONLY
  // meeting_type that means onboarding; every other real value in use
  // ("STE DFU", "STE DEMO", "OSBO DEMO") is confirmed NOT onboarding.
  // This closes off an entire class of bug that's hit this project
  // repeatedly — guessing from title wording alone misfired on the
  // "chase" company-token false match, "Next Steps", "Revisit",
  // "Reconnect", and "Best Practices"/"Click-Around" (a genuine
  // onboarding call and a genuine pre-sale demo that got misclassified
  // in OPPOSITE directions purely from title-text guessing). HubSpot's
  // own meeting objects don't have this field at all — it's Fathom-
  // specific — so they simply fall through to the title-based logic
  // below, completely unaffected by this check.
  if (m && m.meeting_type) {
    var mt = String(m.meeting_type).trim().toLowerCase();
    if (mt === "ste onbd") return true;
    return false; // any other real, known tag is confirmed NOT onboarding
  }

  var title = ((m && (m.meeting_title || m.title)) || "").toLowerCase();
  if (!title) return false;

  // Hard-exclude demos and later touchpoints first, in case a title
  // contains both an onboarding-sounding word and one of these.
  if (title.indexOf("demo") >= 0) return false;
  if (title.indexOf("onesource meeting") >= 0) return false;
  // "click around"/"click-around" confirmed directly as a pre-sale demo,
  // not onboarding — same real distinction Fathom's own meeting_type tag
  // draws (STE DEM vs STE ONBD, see Contact H's case above), but this
  // check matters MORE here: HubSpot's own meeting objects have no
  // meeting_type at all, so for any contact with zero Fathom recordings
  // (confirmed real case: Contact I), this title-level exclusion is
  // the ONLY thing standing between a genuine pre-sale demo and it being
  // wrongly used as the "first onboard date" via the generic name
  // fallback below.
  if (title.indexOf("click around") >= 0 || title.indexOf("click-around") >= 0) return false;
  if (isLaterTouchpointTitle(title)) return false;

  // Explicit onboarding keywords. "training" (whole word) covers both
  // "Training Call" and "OneSource Training - ..."; "onbd" covers "STE
  // ONBD"; "onboard"/"onboarding" (whole word) covers "Onboard Call" —
  // note "onboard" does NOT contain "onbd" as a substring, so this needs
  // its own check rather than relying on the "onbd" match above. "ramp
  // up" covers "OneSource Ramp Up Call" — confirmed a real onboarding
  // call type that HubSpot's title logic didn't recognize at all
  // (Contact I: "0 training call(s)" found despite 3 real Ramp Up
  // calls existing on her profile), forcing a fallback to a year-old demo
  // as a last resort instead.
  if (/\btraining\b/.test(title)) return true;
  if (title.indexOf("onbd") >= 0) return true;
  if (/\bonboard(ing)?\b/.test(title)) return true;
  if (title.indexOf("ramp up") >= 0) return true;

  // Generic "[Client] and [Rep]" fallback for titles with no keyword at
  // all — common in practice (e.g. "Shelley Horb and Sam Stinger Call 3",
  // or "Elia Sanchez and Andrew" where the rep is named by first name only).
  if (contactFirstName) {
    var firstLc = contactFirstName.trim().toLowerCase();
    if (firstLc && title.indexOf(firstLc) >= 0) {
      // 1) Full roster name present (e.g. "Andrew Stein").
      for (var i = 0; i < VALID_REPS.length; i++) {
        if (title.indexOf(VALID_REPS[i].toLowerCase()) >= 0) return true;
      }
      // 2) Unambiguous first name only (e.g. bare "Andrew"). Word-boundary
      // matched so short names like "Ben" don't false-positive inside an
      // unrelated longer word.
      for (var fn in UNAMBIGUOUS_REP_FIRST_NAMES) {
        if (new RegExp("\\b" + fn + "\\b").test(title)) return true;
      }
    }
  }

  return false;
}

// ── Demo / sales-meeting detection ─────────────────────────────────
// The demo (initial sales meeting) identifies the SALESPERSON — via its
// Fathom host (recorded_by), exactly parallel to how the onboarder comes
// from the training call's host. Demos are titled "OneSource Meeting",
// "STE DEMO", or "... Demo ...". A training/onboarding call is never a demo.
function isDemoMeeting(m) {
  // Same authoritative meeting_type check as isOnboardingMeeting above —
  // see that function's comment for the full reasoning. Any tag
  // containing "demo" (covers both "STE DEMO" and "OSBO DEMO") is
  // confirmed a real demo; "STE ONBD" is confirmed NOT a demo. An
  // unrecognized tag (e.g. "STE DFU" — not confirmed as either demo or
  // onboarding) falls through to the title-based check below rather than
  // guessing at what it means.
  if (m && m.meeting_type) {
    var mt = String(m.meeting_type).trim().toLowerCase();
    if (mt.indexOf("demo") >= 0) return true;
    if (mt === "ste onbd") return false;
    // else: unrecognized tag — fall through to title-based check below
  }

  var title = ((m && (m.meeting_title || m.title)) || "").toLowerCase();
  if (!title) return false;

  // A training/onboarding call is not a demo, even if it somehow contained
  // one of the demo words.
  if (/\btraining\b/.test(title) || title.indexOf("onbd") >= 0 || /\bonboard(ing)?\b/.test(title) || title.indexOf("ramp up") >= 0) return false;

  if (title.indexOf("demo") >= 0) return true;
  if (title.indexOf("onesource meeting") >= 0) return true;
  // Confirmed a real pre-sale demo type, distinct from onboarding — see
  // the matching exclusion in isOnboardingMeeting above for the full
  // reasoning. Recognized here too so it can still correctly serve as
  // salesperson-identifying evidence, the same role "demo"/"onesource
  // meeting" already play.
  if (title.indexOf("click around") >= 0 || title.indexOf("click-around") >= 0) return true;

  return false;
}

// ── Attendee-based contact matching ────────────────────────────────
// Confirms a Fathom meeting actually involves this contact, WITHOUT
// relying on Fathom's calendar_invitees[] server filter (which only
// matches formal calendar invitees — it misses people who attended but
// weren't on the invite, e.g. a second person like "Anna" sharing one
// booking with "Liz"). We check, against everything in the meeting object:
//   1. email present anywhere (strongest signal), OR
//   2. first AND last name both present, OR
//   3. first name present AND the company's distinctive words present —
//      onboarding titles include the org (e.g. "... Denny Elwell Company"),
//      so first-name + company is a reliable composite even when the
//      person is listed only by first name with no last name or email.
// Generic company words ("company", "llc", "inc", "group", "realty", etc.)
// are dropped so matching keys on the distinctive part of the org name.
var GENERIC_COMPANY_WORDS = {
  "company":1,"co":1,"llc":1,"inc":1,"inc.":1,"corp":1,"corporation":1,
  "group":1,"the":1,"and":1,"&":1,"of":1,"realty":1,"real":1,"estate":1,
  "commercial":1,"partners":1,"associates":1,"properties":1,"advisors":1,
  "capital":1,"holdings":1,"services":1,"management":1
};

// Personal email providers, excluded from the Fathom company-domain
// filter (see getFathomMeetingRange) since they don't correspond to a
// single real-estate brokerage the way a corporate domain does — querying
// Fathom for "meetings associated with gmail.com" would be meaningless
// (or return an enormous, unrelated set) rather than narrowing anything.
var GENERIC_EMAIL_DOMAINS = {
  "gmail.com":1,"yahoo.com":1,"hotmail.com":1,"outlook.com":1,"icloud.com":1,
  "aol.com":1,"live.com":1,"me.com":1,"msn.com":1,"protonmail.com":1
};

function companyTokens(company) {
  if (!company) return [];
  return String(company).toLowerCase()
    .replace(/[^a-z0-9 ]+/g, " ")
    .split(/\s+/)
    .filter(function(w) { return w && w.length >= 3 && !GENERIC_COMPANY_WORDS[w]; });
}

// Checks Fathom's own CRM-match data (returned when a request passes
// include_crm_matches=true) for a direct link to this exact HubSpot
// contact — by email, the strongest possible signal available, since it
// comes from Fathom's own CRM integration rather than our own text
// matching against the meeting's title/attendee JSON. Only populated when
// the Fathom workspace has its HubSpot CRM connection enabled; harmless
// (just returns false) when it isn't, or when this specific request
// didn't ask for it.
function crmMatchesContact(m, email) {
  var emailLc = (email || "").trim().toLowerCase();
  if (!emailLc || !m || !m.crm_matches || !m.crm_matches.contacts) return false;
  return m.crm_matches.contacts.some(function(c) {
    return c && c.email && String(c.email).trim().toLowerCase() === emailLc;
  });
}

function escapeRegexLiteral(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

// Word-boundary substring check — plain text.indexOf() lets short tokens
// like "reed" or "red" silently match INSIDE unrelated longer words
// ("agreed" contains "reed"; "prepared"/"credentials" contain "red"). A
// meeting whose summary happens to say something like "the client agreed
// to..." would satisfy a raw indexOf("reed") check even though the
// meeting has nothing to do with a contact named Reed. \b enforces this
// only matches the token as its own word. Confirmed real case: Contact A
// was getting cross-matched to an unrelated contact's onboarding call
// (Contact B's, hosted by Ben Ogan) purely through this kind of
// substring collision in contactMatchesMeeting.
function containsWord(text, token) {
  if (!token) return false;
  return new RegExp("\\b" + escapeRegexLiteral(token) + "\\b", "i").test(text);
}

function contactMatchesMeeting(m, firstName, lastName, email, coTokens) {
  // Strongest signal first: Fathom's own CRM match, when available. This
  // can catch cases the text-based checks below miss entirely — e.g. the
  // contact's name/email don't appear anywhere in the meeting's own
  // metadata (they weren't a formal invitee and weren't mentioned by name
  // in the title), but Fathom's CRM integration still correctly linked
  // the call to their HubSpot record.
  if (crmMatchesContact(m, email)) return true;

  var text = JSON.stringify(m).toLowerCase();
  var emailLc = (email || "").trim().toLowerCase();
  var firstLc = (firstName || "").trim().toLowerCase();
  var lastLc  = (lastName || "").trim().toLowerCase();

  if (emailLc && text.indexOf(emailLc) >= 0) return true;
  if (firstLc && lastLc && containsWord(text, firstLc) && containsWord(text, lastLc)) return true;

  if (firstLc && lastLc && coTokens && coTokens.length) {
    // When the company name has MORE THAN ONE distinctive token, require
    // ALL of them to appear (not just one) before counting it as a hit.
    // A single ambiguous token is a real false-positive risk — "NAI Chase
    // Commercial" produces the token "chase", which is also an extremely
    // common first name (and a major bank), so any unrelated meeting that
    // happens to mention someone named Chase, combined with the target
    // contact's own first name appearing ANYWHERE else in that meeting's
    // text, could wrongly match. Requiring every token when there are
    // several (e.g. "nai" AND "chase" both present) makes a coincidental
    // collision far less likely. With only one token available there's no
    // second signal to require, so the single-token check still applies —
    // this residual risk is inherent to companies named after common
    // words, not something a token count alone can fully close.
    //
    // Requires the FULL name (first AND last), not just first name, for
    // the same reason — confirmed real false-positive: Contact G and
    // Contact H both work at "Company A." Contact H's own onboarding
    // call summary happened to mention "Ross" somewhere (a colleague, in
    // passing), and since the company tokens ("stanbrough", "realty")
    // legitimately appeared too (it genuinely was Contact H's real company),
    // the OLD first-name-only check wrongly confirmed the match — pulling
    // an entirely unrelated person's call in as if it were Ross's own.
    // Two coworkers sharing a company and one being mentioned by first
    // name alone in the other's call is a completely ordinary scenario,
    // not a rare edge case — requiring the full name closes this without
    // meaningfully weakening the fallback's original purpose, since a
    // call summary genuinely about a specific person would naturally
    // reference their full name, not just a first name in isolation.
    //
    // Also switched from plain indexOf to word-boundary matching here for
    // the same reason as firstLc/lastLc above — a short company token can
    // collide with substrings inside unrelated words just as easily as a
    // short name can.
    var companyHit = coTokens.length > 1
      ? coTokens.every(function(tok) { return containsWord(text, tok); })
      : containsWord(text, coTokens[0]);
    if (companyHit && containsWord(text, firstLc) && containsWord(text, lastLc)) return true;
  }
  return false;
}

// ── HubSpot ───────────────────────────────────────────────────────
function hsFetch(path, method, payload) {
  var token = getProp("HUBSPOT_TOKEN");
  var opts = {
    method: method || "get",
    headers: { Authorization: "Bearer " + token },
    contentType: "application/json",
    muteHttpExceptions: true
  };
  if (payload) opts.payload = JSON.stringify(payload);
  var resp = UrlFetchApp.fetch("https://api.hubapi.com" + path, opts);
  var code = resp.getResponseCode();
  var body = resp.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error("HubSpot " + code + ": " + body.slice(0, 300));
  }
  return JSON.parse(body);
}

function getHubSpotContact(contactId) {
  var props = "firstname,lastname,email,company,hubspot_owner_id";
  var data = hsFetch("/crm/v3/objects/contacts/" + contactId + "?properties=" + props, "get");
  var p = data.properties || {};
  return {
    firstName: p.firstname || "",
    lastName: p.lastname || "",
    email: p.email || "",
    company: p.company || "",
    ownerId: p.hubspot_owner_id || ""
  };
}

// Fetches the specific company HubSpot's own UI marks as "Primary" for
// this contact — NOT just whichever company happens to be associated.
// Fixed a real bug: a contact linked to two companies (e.g. their actual
// employer, correctly marked Primary in HubSpot's UI, plus a second,
// unrelated company) was pulling the WRONG one, because the old approach
// relied on the legacy `associatedcompanyid` contact property, which does
// not reliably reflect HubSpot's own Primary designation once a contact
// has more than one associated company. The v4 associations endpoint
// returns each associated company along with its actual association
// type — HubSpot's documented standard is associationTypeId 1, labeled
// "Primary" — so this reads that directly rather than trusting a legacy
// single-value field that was never designed for the multi-company case.
function getPrimaryCompanyId(contactId) {
  if (!contactId) return "";
  try {
    var assoc = hsFetch("/crm/v4/objects/contacts/" + contactId + "/associations/companies", "get");
    var results = (assoc && assoc.results) || [];
    for (var i = 0; i < results.length; i++) {
      var types = results[i].associationTypes || [];
      var isPrimary = types.some(function(t) { return t.typeId === 1 || t.label === "Primary"; });
      if (isPrimary) return String(results[i].toObjectId);
    }
    // No association explicitly marked Primary (e.g. only one company is
    // associated at all, which HubSpot doesn't always bother labeling) —
    // fall back to the first result rather than returning nothing.
    if (results.length) return String(results[0].toObjectId);
    return "";
  } catch (e) {
    traceLog("Company", "Could not fetch company associations: " + e.message, "warn");
    return "";
  }
}

// Fetches a company record's name by ID. Used as the fallback when a
// contact's flat `company` property is empty but they're linked to a
// real company record.
function getCompanyName(companyId) {
  if (!companyId) return "";
  try {
    var data = hsFetch("/crm/v3/objects/companies/" + companyId + "?properties=name", "get");
    return (data.properties && data.properties.name || "").trim();
  } catch (e) {
    return "";
  }
}

// Smart company resolution: prefer the PRIMARY associated company
// record — the real, actively-maintained HubSpot association, visible
// in the UI as "Primary" — over the flat `company` text property.
//
// This priority order is deliberate, not the more obvious "flat field
// first": the flat property is a simple, unvalidated text field that can
// go stale independently of a contact's real company associations (a
// past employer, a typo, old imported data) with nothing keeping it in
// sync. Confirmed directly against a real case — a contact's flat
// company property held a stale, incorrect value while their actual,
// current company association correctly showed a different company as
// Primary — and since the flat field wasn't blank, the old logic never
// even got a chance to check the real association at all.
//
// Only falls back to the flat property when there's genuinely no
// company association to check — some contacts (e.g. early-stage
// prospects) may have a company name typed into this field with no
// formal company record created yet, and that's still worth using
// rather than leaving the organization blank.
function resolveCompanyName(contact, contactId) {
  var primaryCompanyId = getPrimaryCompanyId(contactId);
  if (primaryCompanyId) {
    var fromAssoc = getCompanyName(primaryCompanyId);
    if (fromAssoc) {
      traceLog("Company", 'Using PRIMARY associated company record: "' + fromAssoc + '"', "ok");
      return fromAssoc;
    }
  }

  var flat = (contact.company || "").trim();
  if (flat && flat.toLowerCase() !== "unassigned") {
    traceLog("Company", 'No usable company association found — falling back to flat "company" property: "' + flat + '"', "warn");
    return flat;
  }

  traceLog("Company", "No company association and no flat property either — organization will be blank", "fail");
  return "";
}

// Smart name resolution. Some contacts have the company name jammed into
// the first/last name fields — e.g. firstname "Jeffery",
// lastname "Realty - Kaitlyn Mancini" for a person whose real name is
// Kaitlyn Mancini at "Jeffery Realty". When the company name appears
// inside the assembled name, strip it out; if that leaves a plausible
// full name, use it. Otherwise try to reconstruct from a dotted email
// local-part (jane.doe@ -> Jane Doe).
function resolveContactName(contact, companyName) {
  var first = (contact.firstName || "").trim();
  var last  = (contact.lastName || "").trim();
  var email = (contact.email || "").trim();
  var assembled = (first + " " + last).trim();

  var cleaned = assembled;
  var strippedCompany = false;
  if (companyName) {
    var re = new RegExp(companyName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "ig");
    var before = cleaned;
    cleaned = assembled.replace(re, "")
                       .replace(/^[\s\-–—,|]+|[\s\-–—,|]+$/g, "") // trim leftover separators
                       .replace(/\s{2,}/g, " ")                    // collapse doubled spaces
                       .trim();
    strippedCompany = (cleaned !== before);
  }

  // If stripping the company left a plausible full name, use it.
  if (cleaned && cleaned.split(/\s+/).length >= 2) {
    if (strippedCompany) {
      traceLog("Name", 'Raw name "' + assembled + '" contained the company name — stripped to "' + cleaned + '"', "warn");
    } else {
      traceLog("Name", 'Using HubSpot first/last name as-is: "' + cleaned + '"', "ok");
    }
    return cleaned;
  }

  // Otherwise try to reconstruct from a dotted email local-part. Can't
  // reliably split something like "kmancini", so this only fires when
  // the local-part is delimiter-separated.
  if (email) {
    var local = email.split("@")[0];
    if (local.indexOf(".") > 0) {
      var parts = local.split(".").map(function(pp) {
        return pp ? pp.charAt(0).toUpperCase() + pp.slice(1) : "";
      });
      var fromEmail = parts.join(" ").trim();
      if (fromEmail.split(/\s+/).length >= 2) {
        traceLog("Name", 'Name fields unusable ("' + assembled + '") — reconstructed from email local-part: "' + fromEmail + '"', "warn");
        return fromEmail;
      }
    }
  }

  // Last resort: whatever we cleaned, else the raw assembled name, else email.
  var lastResort = cleaned || assembled || email;
  traceLog("Name", 'Could not confidently resolve a clean name — using best-effort value: "' + lastResort + '"', "fail");
  return lastResort;
}

function getOwnerName(ownerId) {
  if (!ownerId) return "";
  try {
    var data = hsFetch("/crm/v3/owners/" + ownerId, "get");
    var name = [data.firstName, data.lastName].filter(function(x){return x;}).join(" ");
    return name || data.email || "";
  } catch (e) {
    return "";
  }
}

// Pulls the salesperson name out of a deal's Source field (e.g.
// "Referral - Sam Absalom"), but ONLY if that name is a valid rep AND the
// source pattern doesn't specifically indicate a cold-call booking.
// "Cold Call - X" names the BOOKER who set the appointment, not
// necessarily the salesperson who actually ran the demo and closed the
// deal — even when X happens to be a valid roster name (some reps also
// do their own cold-calling). Confirmed wrong in practice: a deal's
// Source read "Cold Call - Ben Ogan" for a contact whose real salesperson
// (per Fathom's actual demo recording) was Brock Baker — Ben Ogan was
// just who booked the meeting, not who ran it. So "Cold Call - X" is
// never trusted as salesperson evidence, regardless of whether X is on
// the roster; non-cold-call sources (e.g. a plain rep name, a referral
// source) still get checked against the roster as before.
function salespersonFromSource(source) {
  if (!source) return "";
  if (/\bcold call\b/i.test(source)) return "";
  return validRep(source); // validRep does a substring match against the roster
}

// Given a list of candidate meetings ALREADY SORTED in the desired scan
// order (earliest-first), returns the first one whose owner validates
// against the roster — instead of only ever checking index 0 and giving up
// the moment that one candidate happens to be a non-roster / inactive
// person (e.g. an inactive rep like Caleb Duarte sitting in on a call
// before the real onboarder joined). Every skipped candidate is still
// logged so the trace shows exactly which ones were passed over and why.
// Returns null if NONE of the candidates validate.
function firstValidRepMeeting(sortedMeetings, roleLabel) {
  for (var i = 0; i < sortedMeetings.length; i++) {
    var rawOwner = getOwnerName(sortedMeetings[i]._owner);
    var valid = validRep(rawOwner);
    if (valid) {
      traceLog("HubSpot Meetings→" + roleLabel, 'Meeting "' + (sortedMeetings[i].meeting_title || "(untitled)") + '" owned by "' + rawOwner + '" — valid rep ✓ SELECTED (candidate ' + (i + 1) + ' of ' + sortedMeetings.length + ')', "ok");
      return { meeting: sortedMeetings[i], name: valid, index: i };
    } else if (rawOwner) {
      traceLog("HubSpot Meetings→" + roleLabel, 'Meeting "' + (sortedMeetings[i].meeting_title || "(untitled)") + '" owner "' + rawOwner + '" is not on the valid-rep roster — skipping to next candidate', "warn");
    }
  }
  return null;
}

// Fast, reliable rep resolution straight from the contact's Meetings tab.
// The demo meeting's host = salesperson; the training meeting's host =
// onboarder. We read the meeting OWNER, validated against the roster, and
// take the EARLIEST demo / EARLIEST training WHOSE OWNER ACTUALLY
// VALIDATES — not just the single earliest meeting regardless of owner —
// via firstValidRepMeeting() above. This is a couple of quick HubSpot calls
// (no slow Fathom scan) and has matched the correct rep in every observed
// case where a valid-rep meeting exists among the candidates. Returns
// { salesperson, onboarder } (either may be "").
function getMeetingRepsFromHubSpot(contactId, contactFullName) {
  var out = { salesperson: "", onboarder: "", onboardingScheduledFuture: false, onboardingScheduledDateIso: "",
              onboarderFromPastDemo: false, pastDemoDateIso: "", candidateOnboardingDateIso: "",
              candidateLastOnboardingDateIso: "", knownPastTrainingCount: 0 };
  if (!contactId) return out;

  var contactFirstName = (contactFullName || "").trim().split(/\s+/)[0] || "";

  var assoc;
  try {
    assoc = hsFetch("/crm/v3/objects/contacts/" + contactId + "/associations/meetings", "get");
  } catch (e) {
    traceLog("HubSpot Meetings", "Could not fetch meetings association: " + e.message, "fail");
    return out;
  }

  var results = (assoc && assoc.results) || [];
  if (!results.length) {
    traceLog("HubSpot Meetings", "No meetings found on this contact's HubSpot profile", "warn");
    return out;
  }

  var inputs = [];
  results.forEach(function(r) {
    var id = r.toObjectId || r.id;
    if (id) inputs.push({ id: String(id) });
  });
  if (!inputs.length) return out;

  var batch;
  try {
    batch = hsFetch("/crm/v3/objects/meetings/batch/read", "post", {
      properties: ["hs_meeting_title", "hs_meeting_start_time", "hubspot_owner_id", "hs_internal_meeting_notes"],
      inputs: inputs
    });
  } catch (e) {
    traceLog("HubSpot Meetings", "Batch meeting read failed: " + e.message, "fail");
    return out;
  }

  var meetings = (batch && batch.results) || [];
  var demos = [], trainings = [], allMeetingObjs = [];
  meetings.forEach(function(mm) {
    var p = mm.properties || {};
    // classifiers read meeting_title/title
    var obj = { meeting_title: p.hs_meeting_title || "", _start: p.hs_meeting_start_time || "", _owner: p.hubspot_owner_id || "", _notes: (p.hs_internal_meeting_notes || "").replace(/<[^>]*>/g, "").trim() };
    allMeetingObjs.push(obj);
    if (isDemoMeeting(obj)) demos.push(obj);
    else if (isOnboardingMeeting(obj, contactFirstName)) trainings.push(obj);
  });

  // Reschedule-artifact filter: when two training-titled meetings share
  // the EXACT same title, that's the confirmed signature of a call that
  // got rescheduled — the original booking (now stale) and the new one it
  // moved to both persist as separate meeting records. Directly confirmed
  // real case: Contact I had two meetings both titled "OneSource
  // Ramp Up Call 2 w/ Contact I" (6/4 and 6/9) — the 6/4 one never
  // actually happened and has NO internal meeting notes at all, while the
  // 6/9 one it was rescheduled to has real, substantive notes ("Email Int
  // worked: Covered everything we needed."). This is NOT a blanket "every
  // training call must have notes" rule — a call that genuinely happened
  // hours ago, before the rep has written it up yet, would also have no
  // notes, and that's a normal, common state, not a red flag. The
  // distinction only applies when there's a same-titled TWIN to compare
  // against: if one copy has notes and the other doesn't, the one without
  // is almost certainly the stale pre-reschedule artifact, so it's
  // dropped from the count/date calculation entirely.
  var notesByTitle = {};
  trainings.forEach(function(t) {
    var key = (t.meeting_title || "").trim().toLowerCase();
    if (!key) return;
    if (!notesByTitle[key]) notesByTitle[key] = { withNotes: false, count: 0 };
    notesByTitle[key].count++;
    if (t._notes) notesByTitle[key].withNotes = true;
  });
  var droppedAsRescheduleArtifact = 0;
  trainings = trainings.filter(function(t) {
    var key = (t.meeting_title || "").trim().toLowerCase();
    var group = key ? notesByTitle[key] : null;
    // Only filter when there's an actual duplicate (count > 1), the group
    // has a mix of with/without notes, and THIS specific one has none.
    if (group && group.count > 1 && group.withNotes && !t._notes) {
      droppedAsRescheduleArtifact++;
      return false;
    }
    return true;
  });
  if (droppedAsRescheduleArtifact > 0) {
    traceLog("HubSpot Meetings", "Dropped " + droppedAsRescheduleArtifact + " training-titled meeting(s) as likely reschedule artifacts — a same-titled twin exists WITH real internal notes, while these had none at all (the confirmed signature of a booking that moved but was never cleaned up)", "warn");
  }
  traceLog("HubSpot Meetings", meetings.length + " meeting(s) on profile — " + demos.length + " demo(s), " + trainings.length + " training call(s)", "info");

  var byStart = function(a, b) {
    return (new Date(a._start).getTime() || 0) - (new Date(b._start).getTime() || 0);
  };

  if (demos.length) {
    demos.sort(byStart);
    var demoPick = firstValidRepMeeting(demos, "Salesperson");
    out.salesperson = demoPick ? demoPick.name : "";
  }
  // Tracks the EXACT meeting object (by reference) that supplied the
  // salesperson above, if any — so that if this same meeting later gets
  // reused as an ONBOARDING stand-in (see the three onboarderFromPastDemo
  // branches below), the salesperson pulled from it can be retracted. A
  // single call can't independently evidence two different things: once
  // a meeting has been reclassified as THE onboarding session, its host
  // is already claimed for that role and can't also stand as separate
  // proof of who the salesperson was. Confirmed as a real bug: Contact J's "CRE OneSource Meeting with Collin" tripped the demo-keyword
  // check (naming Collin Michels as salesperson), then — correctly — got
  // reused as her actual onboarding session once the real training call
  // turned out to be future-dated. Without this retraction, Collin ended
  // up wrongly shown as BOTH the salesperson and the onboarder.
  var salespersonSourceMeeting = demoPick ? demoPick.meeting : null;
  if (trainings.length) {
    trainings.sort(byStart);
    var trainingPick = firstValidRepMeeting(trainings, "Onboarder");
    out.onboarder = trainingPick ? trainingPick.name : "";

    // If the earliest EXPLICITLY-titled training call is scheduled in the
    // FUTURE, it hasn't happened yet — but that doesn't necessarily mean
    // nothing has happened at all. Some reps title an already-completed
    // onboarding session as "OneSource Meeting with [rep]" (classified
    // above as a demo by title alone), not "Training Call" — e.g. a call
    // that's clearly an onboarding conversation by content, just not by
    // title convention. If a PAST demo-titled call exists, treat it as the
    // real (already-happened) onboarding session rather than waiting on a
    // future-dated call that may just be a follow-up or duplicate booking.
    // The "is it scheduled in the future" check is a pure scheduling/date
    // fact about the earliest training call and stays based on trainings[0]
    // regardless of who owns it; separately, when we DO fall back to a
    // past demo as the onboarding stand-in, we iterate through ALL past
    // demos (not just the first one chronologically) to find the first one
    // whose owner actually validates.
    var trainingStart = new Date(trainings[0]._start).getTime();
    var trainingIsFuture = trainings[0]._start && !isNaN(trainingStart) && trainingStart > Date.now();
    if (!trainingIsFuture) {
      // A genuine, already-happened training-titled meeting exists at this
      // date — regardless of whether its LISTED OWNER validated (that's
      // exactly what may be wrong; see firstValidRepMeeting above). This
      // is deliberately just a DATE hint, not a trust decision: it lets
      // resolveFathomAndFinalize narrow Fathom's search window tightly
      // around this date so the scan resolves quickly, while Fathom's own
      // recorded_by still has the final say on WHO actually ran the call.
      out.candidateOnboardingDateIso = isoDate(new Date(trainingStart));

      // ALSO expose the LATEST past training-titled meeting's date, not
      // just the earliest — the Fathom search window used to be built
      // from candidateOnboardingDateIso alone (±5 days), which silently
      // missed later real calls for anyone whose onboarding stretched out
      // longer than 5 days. Confirmed as a real bug: Contact D has 4
      // genuine recorded training calls spanning 6/17 to 7/1 (two weeks),
      // but a ±5-day window anchored on just the earliest (6/17) only
      // reaches 6/22 — missing 3 of his 4 real calls entirely and
      // reporting a count of 1 instead of 4. Passing both dates through
      // lets the window span the full known range instead of just one
      // end of it.
      var pastTrainings = [];
      for (var pti = 0; pti < trainings.length; pti++) {
        var ptStart = new Date(trainings[pti]._start).getTime();
        if (!isNaN(ptStart) && ptStart <= Date.now()) pastTrainings.push(trainings[pti]);
      }

      // Stale-history guard: an old, reused contact record can carry
      // training-titled meetings from a PRIOR, completely separate
      // onboarding cycle (a different company, years earlier) still
      // attached to its HubSpot Meetings tab. Left unguarded, the single
      // OVERALL-earliest one becomes "first onboard date" no matter how
      // old it is. Confirmed real case: Contact A's contact record is from
      // 2023 (an old a prior unrelated company-era cycle) — a training-titled
      // meeting from 2023-05-22 was still on his profile alongside his
      // real, current (Aug 2026) onboarding at his own new shop, and got
      // wrongly reported as his first onboard date. Worse, the resulting
      // ~3-year search window in getFathomMeetingRange (built from this
      // date) also raised the odds of an unrelated meeting getting
      // wrongly matched during the Fathom scan — the wide date range
      // widens the haystack, not just the false "first" date.
      //
      // A real, single onboarding cycle's training calls cluster together
      // within weeks, not years. So rather than trust the absolute
      // earliest, walk backward from the MOST RECENT past training call
      // and keep including older ones only while each stays within
      // STALE_GAP_DAYS of its next-more-recent neighbor. Anything further
      // back than that gap is treated as leftover history from a
      // separate, prior cycle and excluded from the date/count/search-
      // window entirely — genuinely old, unrelated activity shouldn't
      // silently masquerade as part of the current onboarding.
      var recentCycleAnchor = trainings[0]; // fallback if pastTrainings is empty (shouldn't happen inside !trainingIsFuture, but keeps this safe)
      if (pastTrainings.length) {
        pastTrainings.sort(byStart);
        var STALE_GAP_DAYS = 90;
        var currentCycle = [pastTrainings[pastTrainings.length - 1]];
        for (var pci = pastTrainings.length - 2; pci >= 0; pci--) {
          var newerStart = new Date(currentCycle[0]._start).getTime();
          var olderStart = new Date(pastTrainings[pci]._start).getTime();
          var gapDays = (newerStart - olderStart) / (1000 * 60 * 60 * 24);
          if (gapDays <= STALE_GAP_DAYS) {
            currentCycle.unshift(pastTrainings[pci]);
          } else {
            traceLog("HubSpot Meetings", "Excluding " + (pci + 1) + " older training-titled meeting(s) as likely leftover history from a separate, prior onboarding cycle — earliest excluded one (\"" + (pastTrainings[pci].meeting_title || "(untitled)") + "\") on " + isoDate(new Date(olderStart)) + " is " + Math.round(gapDays) + " day(s) older than the next-most-recent real training call (" + STALE_GAP_DAYS + "-day cutoff). Using only the recent cluster for first-onboard-date, count, and the Fathom search window.", "warn");
            break; // everything further back is even older — stop, we found the boundary
          }
        }
        out.candidateOnboardingDateIso = isoDate(new Date(new Date(currentCycle[0]._start).getTime()));
        out.candidateLastOnboardingDateIso = isoDate(new Date(new Date(currentCycle[currentCycle.length - 1]._start).getTime()));
        // Known count of PAST (already-happened) training-titled meetings
        // on HubSpot's own Meetings tab — used downstream as the actual
        // TARGET for the Fathom search, instead of a flat page-count
        // tolerance that's wrong for every contact in one direction or
        // the other (too aggressive for a simple 1-call contact, still
        // not enough for a contact with several scattered real calls).
        // Knowing "there should be exactly N" lets the search stop
        // immediately and confidently once N is found, or keep going with
        // real justification when it isn't — rather than guessing either
        // way. Counts only the recent cluster, not the stale leftovers
        // excluded above.
        out.knownPastTrainingCount = currentCycle.length;
        recentCycleAnchor = currentCycle[0];
      }

      // BUT: an even EARLIER meeting from the SAME rep might be the TRUE
      // first onboarding session, simply with no title at all — the same
      // blind spot the future-training branch below already handles, just
      // not gated on "is trainings[0] still scheduled in the future."
      // Confirmed as a real, separate bug: a contact's earliest classified
      // training call was future-dated at first (correctly triggering the
      // stand-in logic below and finding a real earlier untitled call) —
      // but once enough real time passed that the SAME classified call
      // became past-dated, this branch fired instead, which never checked
      // for an earlier untitled call at all. The result was a fast-path
      // date that was technically real but not the earliest one, AND —
      // since the true earliest call is still invisible to Fathom's own
      // title-based classification too — an expensive, doomed multi-tier
      // Fathom scan that could never find what it was looking for (it hit
      // the full sequential fallback and real rate limits chasing a call
      // Fathom's own classifier can never recognize as "onboarding").
      // Checking for an earlier same-owner stand-in unconditionally, and
      // skipping Fathom entirely when one is found, avoids that cost.
      // Only consider candidates that AREN'T already a legitimately
      // classified demo or training call — a real demo already supplies
      // independent salesperson evidence and must never be repurposed as
      // an onboarding stand-in just because it happens to be earlier;
      // that's exactly the mistake this fix would otherwise make. This is
      // strictly for meetings invisible to classification entirely (e.g.
      // no title at all), which is the actual gap being closed.
      //
      // This check now anchors off the recent cluster's own earliest
      // meeting (recentCycleAnchor), not the raw absolute-earliest
      // training call — otherwise this check would itself go chasing an
      // owner match against a stale, years-old meeting, the exact same
      // failure mode the stale-history guard above exists to prevent.
      var isAlreadyClassified = function(obj) {
        for (var dci = 0; dci < demos.length; dci++) { if (demos[dci] === obj) return true; }
        for (var tci = 0; tci < trainings.length; tci++) { if (trainings[tci] === obj) return true; }
        return false;
      };
      var recentCycleStart = new Date(recentCycleAnchor._start).getTime();
      var earlierSameOwnerPast = [];
      for (var ei = 0; ei < allMeetingObjs.length; ei++) {
        var em = allMeetingObjs[ei];
        var eStart = new Date(em._start).getTime();
        // Restricted to genuinely UNTITLED meetings (no title at all) —
        // NOT merely "unclassified". A meeting that HAS a real title but
        // simply doesn't match our keywords (e.g. "CRE OneSource Reconnect
        // with Andrew Stein") is a real, distinct signal on its own — it
        // could be a sales re-engagement call, a casual check-in, anything
        // — and promoting it to "the onboarding session" purely for being
        // chronologically earlier and sharing an owner is a much weaker
        // basis than genuine title-blindness. Confirmed as a real bug:
        // Contact F had a Feb 2 call titled "Reconnect" (a pre-deal
        // sales touchpoint, months before onboarding even started) that
        // got wrongly promoted over the real June 30 training call this
        // way, reporting a wildly wrong first-onboard date. The untitled
        // case this fix now scopes to is Contact E's — a call with NO
        // title at all, which is invisible to isOnboardingMeeting/
        // isDemoMeeting by construction, not just a title we don't
        // recognize.
        var emTitle = (em.meeting_title || "").trim();
        if (em._owner && em._owner === recentCycleAnchor._owner && !isNaN(eStart) && eStart < recentCycleStart && !emTitle && !isAlreadyClassified(em)) {
          earlierSameOwnerPast.push(em);
        }
      }
      if (earlierSameOwnerPast.length) {
        earlierSameOwnerPast.sort(byStart);
        var earlierPick = firstValidRepMeeting(earlierSameOwnerPast, "Onboarder (earlier same-owner call)");
        if (earlierPick) {
          traceLog("HubSpot Meetings", 'A meeting on ' + isoDate(new Date(new Date(earlierPick.meeting._start).getTime())) + ' (title: "' + (earlierPick.meeting.meeting_title || "(untitled)") + '"), owned by the SAME rep as the current cycle\'s earliest classified training call ("' + (recentCycleAnchor.meeting_title || "(untitled)") + '" on ' + isoDate(new Date(recentCycleStart)) + '), happened earlier — treating that as the real first onboarding session instead. A call can be the real thing even with no title at all.', "warn");
          out.onboarder = earlierPick.name;
          out.onboarderFromPastDemo = true;
          out.pastDemoDateIso = isoDate(new Date(new Date(earlierPick.meeting._start).getTime()));
          if (salespersonSourceMeeting && earlierPick.meeting === salespersonSourceMeeting) {
            traceLog("HubSpot Meetings→Salesperson", 'Retracting salesperson "' + out.salesperson + '" — it came from this SAME meeting, which is now being used as the onboarding session, not independent demo evidence. Falling back to the deal for salesperson instead.', "warn");
            out.salesperson = "";
          }
        }
      }
    }
    if (trainingIsFuture) {
      var pastDemosForFuture = [];
      for (var di = 0; di < demos.length; di++) {
        var dStart = new Date(demos[di]._start).getTime();
        if (!isNaN(dStart) && dStart <= Date.now()) pastDemosForFuture.push(demos[di]);
      }
      if (pastDemosForFuture.length) {
        var pastPick = firstValidRepMeeting(pastDemosForFuture, "Onboarder (past-demo stand-in)");
        if (pastPick) {
          traceLog("HubSpot Meetings", 'Earliest "training" call is scheduled for ' + isoDate(new Date(trainingStart)) + ' (future) — but a past call "' + (pastPick.meeting.meeting_title || "(untitled)") + '" on ' + isoDate(new Date(new Date(pastPick.meeting._start).getTime())) + ' already happened. Treating that as the real onboarding session instead of waiting on the future one.', "warn");
          out.onboarder = pastPick.name;
          out.onboarderFromPastDemo = true;
          out.pastDemoDateIso = isoDate(new Date(new Date(pastPick.meeting._start).getTime()));
          if (salespersonSourceMeeting && pastPick.meeting === salespersonSourceMeeting) {
            traceLog("HubSpot Meetings→Salesperson", 'Retracting salesperson "' + out.salesperson + '" — it came from this SAME meeting, which is now being used as the onboarding session, not independent demo evidence. Falling back to the deal for salesperson instead.', "warn");
            out.salesperson = "";
          }
        }
      } else {
        // Still nothing to stand in for it — but a meeting with NO TITLE
        // AT ALL is invisible to isDemoMeeting/isOnboardingMeeting (both
        // bail out immediately on an empty title string), even though it
        // may be the REAL onboarding session: a rep runs the actual call
        // untitled, then later books a formally-titled follow-up for the
        // same relationship, and this future-titled one is that follow-up
        // — not the original. Same-owner continuity (the earlier meeting
        // is owned by the SAME rep as the future training call) is a much
        // stronger signal than title text alone, so this only fires when
        // title-based classification already came up completely empty.
        // Restricted to genuinely UNTITLED meetings only — see the
        // matching comment on earlierSameOwnerPast above for why a
        // titled-but-unrecognized meeting (e.g. "Reconnect") must NOT be
        // promoted this way, unlike a truly blank title.
        var sameOwnerPast = [];
        for (var ai = 0; ai < allMeetingObjs.length; ai++) {
          var am = allMeetingObjs[ai];
          var aStart = new Date(am._start).getTime();
          var amTitle = (am.meeting_title || "").trim();
          if (am._owner && am._owner === trainings[0]._owner && !isNaN(aStart) && aStart <= Date.now() && !amTitle) {
            sameOwnerPast.push(am);
          }
        }
        if (sameOwnerPast.length) {
          sameOwnerPast.sort(byStart);
          var untitledPick = firstValidRepMeeting(sameOwnerPast, "Onboarder (untitled same-owner stand-in)");
          if (untitledPick) {
            traceLog("HubSpot Meetings", 'Earliest "training" call is scheduled for ' + isoDate(new Date(trainingStart)) + ' (future) — but a meeting on ' + isoDate(new Date(new Date(untitledPick.meeting._start).getTime())) + ' (title: "' + (untitledPick.meeting.meeting_title || "(untitled)") + '"), owned by the SAME rep, already happened. Treating that as the real onboarding session — a call can be the real thing even with no title at all.', "warn");
            out.onboarder = untitledPick.name;
            out.onboarderFromPastDemo = true;
            out.pastDemoDateIso = isoDate(new Date(new Date(untitledPick.meeting._start).getTime()));
            if (salespersonSourceMeeting && untitledPick.meeting === salespersonSourceMeeting) {
              traceLog("HubSpot Meetings→Salesperson", 'Retracting salesperson "' + out.salesperson + '" — it came from this SAME meeting, which is now being used as the onboarding session, not independent demo evidence. Falling back to the deal for salesperson instead.', "warn");
              out.salesperson = "";
            }
          } else {
            out.onboardingScheduledFuture = true;
            out.onboardingScheduledDateIso = isoDate(new Date(trainingStart));
            traceLog("HubSpot Meetings", "Earliest training call is scheduled for " + out.onboardingScheduledDateIso + " — in the future, hasn't happened yet, and no past call (titled or not) from the same rep exists to use instead. Skipping the Fathom search for onboard dates.", "warn");
          }
        } else {
          out.onboardingScheduledFuture = true;
          out.onboardingScheduledDateIso = isoDate(new Date(trainingStart));
          traceLog("HubSpot Meetings", "Earliest training call is scheduled for " + out.onboardingScheduledDateIso + " — in the future, hasn't happened yet, and no past call exists to use instead. Skipping the Fathom search for onboard dates.", "warn");
        }
      }
    }
  } else if (demos.length) {
    // No explicitly-titled training call exists AT ALL — if there's only
    // ONE relevant call (common for lighter-touch accounts, e.g. an
    // intern), it likely combined the demo and onboarding into one
    // session. Iterate through every PAST demo (not just the earliest one)
    // to find the first whose owner actually validates, rather than
    // checking only the first past candidate and giving up if it happens
    // to be a non-roster person.
    var pastDemos = [];
    for (var si = 0; si < demos.length; si++) {
      var sStart = new Date(demos[si]._start).getTime();
      if (!isNaN(sStart) && sStart <= Date.now()) pastDemos.push(demos[si]);
    }
    if (pastDemos.length) {
      var soloPick = firstValidRepMeeting(pastDemos, "Onboarder (combined demo+onboarding)");
      if (soloPick) {
        traceLog("HubSpot Meetings", 'No dedicated training call found — using "' + (soloPick.meeting.meeting_title || "(untitled)") + '" as a combined demo+onboarding session.', "warn");
        out.onboarder = soloPick.name;
        out.onboarderFromPastDemo = true;
        out.pastDemoDateIso = isoDate(new Date(new Date(soloPick.meeting._start).getTime()));
        if (salespersonSourceMeeting && soloPick.meeting === salespersonSourceMeeting) {
          traceLog("HubSpot Meetings→Salesperson", 'Retracting salesperson "' + out.salesperson + '" — it came from this SAME meeting, which is now being used as the onboarding session, not independent demo evidence. Falling back to the deal for salesperson instead.', "warn");
          out.salesperson = "";
        }
      }
    }
  }
  return out;
}

function getContactDeal(contactId) {
  var assoc;
  try {
    assoc = hsFetch("/crm/v3/objects/contacts/" + contactId + "/associations/deals", "get");
  } catch (e) {
    traceLog("Deal", "Could not fetch associated deals: " + e.message, "fail");
    return { salesperson: "", source: "", createdDateIso: "", suggestedUserTypeFromDealSize: "", dealProducts: "" };
  }
  var results = assoc.results || [];
  if (!results.length) {
    traceLog("Deal", "No associated deal found on this contact", "warn");
    return { salesperson: "", source: "", createdDateIso: "", suggestedUserTypeFromDealSize: "", dealProducts: "" };
  }

  var dealId = results[results.length - 1].id || results[results.length - 1].toObjectId;
  if (!dealId) return { salesperson: "", source: "", createdDateIso: "", suggestedUserTypeFromDealSize: "", dealProducts: "" };

  var deal = hsFetch("/crm/v3/objects/deals/" + dealId + "?properties=hubspot_owner_id,source,createdate,deal_size,product_s_", "get");
  var dp = deal.properties || {};
  traceLog("Deal", 'Found deal ' + dealId + ' — Source: "' + (dp.source || "(none)") + '"', "info");

  // Prefer the salesperson named in the Source field; only fall back to the
  // deal owner's name if Source doesn't name a recognizable team member OR
  // if it's a "Cold Call - X" pattern (names the booker, not the closer —
  // rejected regardless of whether X is on the roster; see
  // salespersonFromSource).
  var salesperson = salespersonFromSource(dp.source);
  if (salesperson) {
    traceLog("Deal→Salesperson", 'Source names a valid rep: "' + salesperson + '" (usable only as a last-resort fallback)', "ok");
  } else {
    if (dp.source && /\bcold call\b/i.test(dp.source)) {
      traceLog("Deal→Salesperson", 'Source "' + dp.source + '" is a cold-call booking — rejected regardless of whether the named person is a valid rep (that field names the booker, not necessarily who closed the deal)', "warn");
    } else if (dp.source) {
      traceLog("Deal→Salesperson", 'Source "' + dp.source + '" is not a valid rep — rejected', "warn");
    }
    salesperson = getOwnerName(dp.hubspot_owner_id);
    if (salesperson) traceLog("Deal→Salesperson", 'Falling back to deal owner "' + salesperson + '" — NOT roster-validated at this stage', "warn");
  }

  // Primary User Type signal — a real category your own sales team
  // recorded directly on the deal ("Deal Size": Individual, Team, or
  // Company), not an inference from call attendees. "Company" is mapped
  // to "Team" here — Lupe's own User Type dropdown only has two options,
  // and Company is the closest fit to Team of the two.
  var dealSizeRaw = String(dp.deal_size || "").trim();
  var dealSizeLc = dealSizeRaw.toLowerCase();
  var suggestedUserTypeFromDealSize = "";
  if (dealSizeLc === "individual") suggestedUserTypeFromDealSize = "Individual";
  else if (dealSizeLc === "team" || dealSizeLc === "company") suggestedUserTypeFromDealSize = "Team";
  if (suggestedUserTypeFromDealSize) {
    traceLog("Deal→User Type", '"Deal Size" = "' + dealSizeRaw + '" — suggesting "' + suggestedUserTypeFromDealSize + '"' + (dealSizeLc === "company" ? " (Company mapped to Team)" : "") + " (editable)", "ok");
  }

  return { salesperson: salesperson, source: dp.source || "", createdDateIso: dp.createdate ? isoDate(new Date(dp.createdate)) : "",
           suggestedUserTypeFromDealSize: suggestedUserTypeFromDealSize, dealProducts: dp.product_s_ || "" };
}

// ── Owner lookup by name (reverse of getOwnerName) ─────────────────
// Fetches all HubSpot owners and finds the one whose full name matches.
// Used to turn a Fathom-identified onboarder's NAME into the owner ID
// HubSpot needs for updating hubspot_owner_id or assigning a task.
function findOwnerIdByName(fullName) {
  if (!fullName) return null;
  var nameLc = fullName.trim().toLowerCase();
  try {
    var data = hsFetch("/crm/v3/owners?limit=100", "get");
    var owners = data.results || [];
    for (var i = 0; i < owners.length; i++) {
      var o = owners[i];
      var oName = [o.firstName, o.lastName].filter(function(x){return x;}).join(" ").trim().toLowerCase();
      if (oName === nameLc) return String(o.id);
    }
    // fallback: loose match (handles minor formatting differences)
    for (var j = 0; j < owners.length; j++) {
      var o2 = owners[j];
      var oName2 = [o2.firstName, o2.lastName].filter(function(x){return x;}).join(" ").trim().toLowerCase();
      if (oName2 && (oName2.indexOf(nameLc) >= 0 || nameLc.indexOf(oName2) >= 0)) return String(o2.id);
    }
  } catch (e) {
    // fall through
  }
  return null;
}

// ── Called from the frontend: update contact owner to the onboarder ──
// SOP Task 3: contact owner should always be whoever conducted the
// onboarding call. Returns {ok, error?} so the frontend can show a
// clear per-action result rather than a single all-or-nothing status.
function updateContactOwnerToOnboarder(contactId, onboarderName) {
  if (!contactId) return { ok: false, error: "Missing contact ID." };
  if (!onboarderName) return { ok: false, error: "No onboarder identified — nothing to set." };

  var ownerId = findOwnerIdByName(onboarderName);
  if (!ownerId) return { ok: false, error: 'Could not find a HubSpot owner matching "' + onboarderName + '".' };

  try {
    hsFetch("/crm/v3/objects/contacts/" + contactId, "patch", {
      properties: { hubspot_owner_id: ownerId }
    });
    return { ok: true, ownerId: ownerId };
  } catch (e) {
    return { ok: false, error: e.message || String(e) };
  }
}

// ── Called from the frontend: create the thank-you card task ──────
// SOP Task 4: always assigned to Peter Billing, no further action needed
// beyond creating it. Associates the task with the contact so it shows
// up on their record.
function createThankYouTask(contactId, contactName) {
  if (!contactId) return { ok: false, error: "Missing contact ID." };

  try {
    // SOP Task 4: due Day 1 — i.e. the day this is being processed, not
    // some arbitrary number of days out.
    var dueDate = new Date();

    var task = hsFetch("/crm/v3/objects/tasks", "post", {
      properties: {
        hs_task_subject: "Send thank-you card" + (contactName ? " — " + contactName : ""),
        hs_task_status: "NOT_STARTED",
        hs_task_type: "TODO",
        hs_timestamp: dueDate.toISOString(),
        hubspot_owner_id: PETER_BILLING_OWNER_ID
      }
    });

    var taskId = task.id;
    if (!taskId) return { ok: false, error: "Task created but no ID returned." };

    // Associate the task with the contact (best-effort — task still exists
    // even if this association call fails for some reason).
    try {
      hsFetch("/crm/v4/objects/tasks/" + taskId + "/associations/default/contacts/" + contactId, "put", {});
    } catch (assocErr) {
      return { ok: true, taskId: taskId, warning: "Task created but couldn't associate it with the contact: " + assocErr.message };
    }

    return { ok: true, taskId: taskId };
  } catch (e) {
    return { ok: false, error: e.message || String(e) };
  }
}

// ── Generic workflow task creation ─────────────────────────────────
// Used by any of the numbered SOP tasks (Suite Tasks 5-15, etc.) that
// the person has toggled "Yes" on in the workflow list. Unlike the
// thank-you task (which always has the same title and always goes to
// Peter), this handles arbitrary title/type/due-date/assignee.
//
// HubSpot's hs_task_type property only supports a small fixed set —
// mapping our SOP's four labels: Call -> CALL, Email -> EMAIL,
// everything else (General Task, Text) -> TODO, with the true label
// kept in the title itself since HubSpot has no native "Text" task type.
function mapTaskType(sopType) {
  if (sopType === "Call") return "CALL";
  if (sopType === "Email") return "EMAIL";
  return "TODO"; // General Task, Text
}

// Sends the Suite/Prospects post-onboard email directly from Gmail,
// requiring a human to have reviewed and clicked Send in Lupe's own UI
// first — this function only ever runs after that confirmation, never on
// its own. IMPORTANT: correct per-onboarder attribution (sending as
// whichever person is actually using Lupe, not a single fixed account)
// requires the web app deployment's "Execute as" setting to be changed
// to "User accessing the web app" — with the default "Execute as: Me"
// setting, every send here would come from whoever originally deployed
// the script, regardless of who clicked Send in the UI. Each person also
// needs to go through a one-time Google consent screen the first time
// they use this, under that deployment setting.
// Derives a proper "First Last" display name for whoever's currently
// executing the script (the accessing user, once the deployment runs as
// "User accessing the web app") — without this, GmailApp.sendEmail falls
// back to showing the raw email local-part as the sender name (e.g.
// "andrew.yip" instead of "Andrew Yip"), which looks unpolished and can
// read as spam-like to a recipient. Checks the existing onboarder roster
// first, since it has the authoritative, exact spelling (handles
// hyphens/apostrophes correctly where a naive derivation might not); this
// is the ONLY real fallback for someone not on that roster at all — e.g.
// the script's own deployer, who manages/deploys Lupe but isn't
// themselves an assignable onboarder.
function deriveSenderDisplayName() {
  var email = Session.getActiveUser().getEmail();
  if (!email) return "";
  for (var name in REP_EMAILS) {
    if (REP_EMAILS[name].toLowerCase() === email.toLowerCase()) return name;
  }
  var localPart = email.split("@")[0];
  return localPart.split(/[._-]+/).map(function(part) {
    return part ? (part.charAt(0).toUpperCase() + part.slice(1).toLowerCase()) : "";
  }).filter(function(p) { return p; }).join(" ");
}

function sendPostOnboardEmail(toEmail, subject, htmlBody, attachmentKeys, ccEmail, bccAddress) {
  if (!toEmail) return { ok: false, error: "Missing recipient email." };
  if (!subject) return { ok: false, error: "Missing subject." };
  if (!htmlBody) return { ok: false, error: "Missing email body." };
  try {
    var senderName = deriveSenderDisplayName();
    var options = { htmlBody: htmlBody };
    if (senderName) options.name = senderName;
    if (attachmentKeys && attachmentKeys.length) {
      options.attachments = getProspectsWelcomeAttachments(attachmentKeys);
    }
    // Optional CC — GmailApp's own cc option accepts a comma-separated
    // string for multiple addresses natively, so this just passes
    // whatever was entered straight through rather than parsing it.
    if (ccEmail) options.cc = ccEmail;
    // BCC to the sending onboarder's own HubSpot logging address — this is
    // what makes HubSpot log the send to the contact's timeline at all.
    // Required on the Index.html side before Send is even clickable, but
    // this function stays defensive and simply omits bcc if it's missing
    // rather than failing the send outright.
    if (bccAddress) options.bcc = bccAddress;
    GmailApp.sendEmail(toEmail, subject, "", options);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message || String(e) };
  }
}

// GmailApp.sendEmail does NOT automatically append the sender's Gmail
// signature the way Gmail's own compose window does — confirmed
// directly, this is a real, long-standing Apps Script limitation (there's
// no supported way to read a user's actual Gmail signature setting via
// Apps Script at all). The workaround here: each onboarder pastes their
// own signature into Lupe once, and it's remembered from then on.
//
// PropertiesService.getUserProperties() is scoped per ACCESSING person
// once the deployment runs as "User accessing the web app" — the exact
// same setting sendPostOnboardEmail already needs for correct per-person
// send attribution, so this piggybacks on that same requirement rather
// than needing anything additional. Under the default "Execute as: Me"
// setting, this would incorrectly be scoped to the script owner instead
// of whoever's actually using Lupe.
function getMySignature() {
  return PropertiesService.getUserProperties().getProperty("emailSignatureHtml") || "";
}

function saveMySignature(signatureHtml) {
  if (!signatureHtml) return { ok: false, error: "Signature can't be empty." };
  PropertiesService.getUserProperties().setProperty("emailSignatureHtml", signatureHtml);
  return { ok: true };
}

// HubSpot BCC logging address — same one-time-setup pattern and same
// PropertiesService USER-scoped storage as the signature above, just a
// different property key, so each onboarder's BCC address is private to
// their own login exactly like their signature is. Unlike the signature,
// this is REQUIRED before sending (enforced in Index.html), since skipping
// it means the send goes out completely untracked in HubSpot rather than
// just missing a nice-to-have sign-off.
function getMyBccAddress() {
  return PropertiesService.getUserProperties().getProperty("bccAddress") || "";
}

function saveMyBccAddress(address) {
  if (!address) return { ok: false, error: "BCC address can't be empty." };
  PropertiesService.getUserProperties().setProperty("bccAddress", address);
  return { ok: true };
}

// dueDateIso: "YYYY-MM-DD" string, already computed and confirmed by the
// frontend (weekend-shifted, etc.) — this function trusts that and just
// creates the task, it does not recompute or validate the date logic.
//
// assigneeName: a person's full name (onboarder or salesperson) OR null
// if using Peter directly. If a name is given but no matching HubSpot
// owner can be found, the task is still created but left UNASSIGNED,
// with a warning returned — this lets testing proceed safely even when
// Fathom/deal data is missing, rather than failing outright.
function createWorkflowTask(contactId, title, sopType, dueDateIso, assigneeKind, onboarderName, salespersonName, notes) {
  if (!contactId) return { ok: false, error: "Missing contact ID." };
  if (!dueDateIso) return { ok: false, error: "Missing due date." };

  var ownerId = null;
  var warning = null;

  if (assigneeKind === "peter" || assigneeKind === "Peter Billing") {
    ownerId = PETER_BILLING_OWNER_ID;
  } else if (assigneeKind === "onboarder") {
    ownerId = findOwnerIdByName(onboarderName);
    if (!ownerId) warning = 'No HubSpot owner found for onboarder "' + (onboarderName || "(blank)") + '" — task created unassigned.';
  } else if (assigneeKind === "salesperson") {
    ownerId = findOwnerIdByName(salespersonName);
    if (!ownerId) warning = 'No HubSpot owner found for salesperson "' + (salespersonName || "(blank)") + '" — task created unassigned.';
  } else if (assigneeKind) {
    // Manual per-task override from Lupe's assignee dropdown — a direct
    // rep name rather than one of the three SOP-driven kinds above.
    ownerId = findOwnerIdByName(assigneeKind);
    if (!ownerId) warning = 'No HubSpot owner found for "' + assigneeKind + '" — task created unassigned.';
  }

  try {
    var props = {
      hs_task_subject: title,
      hs_task_status: "NOT_STARTED",
      hs_task_type: mapTaskType(sopType),
      hs_timestamp: new Date(dueDateIso + "T08:00:00").toISOString()
    };
    if (ownerId) props.hubspot_owner_id = ownerId;
    // hs_task_body is HubSpot's actual "Task Notes" field, shown on the
    // task detail view — deliberately never rendered or editable in
    // Lupe's own UI (that's SOP-fixed content: call scripts, thank-you
    // card options), it only ever appears on the real HubSpot task itself.
    if (notes) props.hs_task_body = notes;

    var task = hsFetch("/crm/v3/objects/tasks", "post", { properties: props });
    var taskId = task.id;
    if (!taskId) return { ok: false, error: "Task created but no ID returned." };

    try {
      hsFetch("/crm/v4/objects/tasks/" + taskId + "/associations/default/contacts/" + contactId, "put", {});
    } catch (assocErr) {
      return { ok: true, taskId: taskId, warning: (warning ? warning + " " : "") + "Also couldn't associate with the contact: " + assocErr.message };
    }

    return { ok: true, taskId: taskId, warning: warning };
  } catch (e) {
    return { ok: false, error: e.message || String(e) };
  }
}

// ── DEBUG: test task creation on ONE real contact, in isolation ───
// Run this directly from the Apps Script editor to check whether task
// creation actually works with your current HubSpot token/scopes —
// without touching the sheet or going through the full app flow.
// Result pops up as an alert (switch to the Sheet tab to see it).
function debugCreateTestTask() {
  var contactIdOrUrl = "PASTE A REAL CONTACT URL OR ID HERE";
  var contactName = "Test Contact"; // just used in the task title

  var ui = SpreadsheetApp.getUi();
  var contactId = extractContactId(contactIdOrUrl) || contactIdOrUrl;

  if (!contactId) {
    ui.alert("Couldn't resolve a contact ID from: " + contactIdOrUrl);
    return;
  }

  var result = createThankYouTask(contactId, contactName);

  var msg;
  if (result.ok) {
    msg = "✅ SUCCESS\n\n"
      + "Task created with ID: " + result.taskId + "\n"
      + "Assigned to: Peter Billing (owner ID " + PETER_BILLING_OWNER_ID + ")\n"
      + "Associated with contact: " + contactId
      + (result.warning ? "\n\n⚠ " + result.warning : "")
      + "\n\nGo check that contact's Tasks tab in HubSpot to confirm it shows up.";
  } else {
    msg = "❌ FAILED\n\n"
      + "Contact ID: " + contactId + "\n"
      + "Error: " + result.error
      + "\n\nIf this mentions a missing scope, that tells us exactly what to add to the private app.";
  }

  ui.alert("Task Creation Test", msg, ui.ButtonSet.OK);
}

// ── Workflow enrollment (real HubSpot workflows, not Sales sequences) ──
// Enrolls a contact into one of the two SOP-required post-onboard email
// workflows — replacing what used to be a manual "Add to Post Onboard
// Workflow" reminder task with Lupe actually doing it.
//
// Two real technical quirks made this necessary to build carefully,
// confirmed directly through testing rather than assumed from docs:
//
// 1. HubSpot's contact-enrollment endpoint is still the older v2 API —
//    there is no newer, documented v4 equivalent for enrolling a CONTACT
//    into an existing workflow (v4 covers creating/reading/updating the
//    workflow's own structure, not enrollment into it).
// 2. That v2 endpoint does not accept the newer-style flow ID shown in
//    HubSpot's own UI URLs (confirmed directly: a real, valid flow ID
//    returns a genuine 404 "resource not found" from this endpoint).
//    It requires a separate, officially-documented mapping step first —
//    POST /automation/v4/workflow-id-mappings/batch/read — to convert
//    the newer flowId into the older workflowId the v2 endpoint expects.
//
// contactEmail: the contact's email address (the v2 endpoint enrolls by
// email, not by contact ID).
// flowId: the newer-style ID from the workflow's own HubSpot URL.
//
// Returns { ok: true } on a 204 (HubSpot's documented success response —
// an empty body, not JSON, which is why this uses a raw UrlFetchApp call
// rather than the shared hsFetch() helper; hsFetch() unconditionally
// tries to JSON-parse the response body, which would incorrectly throw
// on this endpoint's genuine empty-body success case).
// Returns { ok: false, error: "..." } otherwise — including if the
// mapping step itself fails, or if the contact's email isn't found.
//
// IMPORTANT CAVEAT, confirmed directly during testing and worth knowing:
// other HubSpot API users have found this v2 endpoint can return 204
// even when nothing actually happened (e.g. re-enrolling someone already
// enrolled in a workflow that doesn't allow re-enrollment) — a 204 here
// means "HubSpot accepted the request," not a guarantee the contact is
// now genuinely enrolled. Worth spot-checking HubSpot's own workflow
// enrollment history occasionally, especially for a contact who may
// have already been through this same workflow before.
function enrollContactInWorkflow(contactEmail, flowId) {
  if (!contactEmail) return { ok: false, error: "Missing contact email." };
  if (!flowId) return { ok: false, error: "Missing flow ID." };
  var token = getProp("HUBSPOT_TOKEN");

  // Step 1: convert the newer flowId into the older workflowId the v2
  // enrollment endpoint actually expects.
  var mapResp = UrlFetchApp.fetch("https://api.hubapi.com/automation/v4/workflow-id-mappings/batch/read", {
    method: "post",
    headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
    payload: JSON.stringify({ inputs: [{ flowMigrationStatuses: flowId, type: "FLOW_ID" }] }),
    muteHttpExceptions: true
  });
  var mapCode = mapResp.getResponseCode();
  if (mapCode < 200 || mapCode >= 300) {
    return { ok: false, error: "Workflow ID mapping failed (HTTP " + mapCode + "): " + mapResp.getContentText().slice(0, 200) };
  }
  var mapped;
  try {
    mapped = JSON.parse(mapResp.getContentText());
  } catch (e) {
    return { ok: false, error: "Workflow ID mapping returned unreadable data." };
  }
  var mappedId = (mapped.results && mapped.results.length) ? mapped.results[0].workflowId : null;
  if (!mappedId) {
    return { ok: false, error: "Could not find a mapped workflow ID for flow " + flowId + "." };
  }

  // Step 2: enroll the contact using the mapped ID.
  var enrollUrl = "https://api.hubapi.com/automation/v2/workflows/" + mappedId + "/enrollments/contacts/" + encodeURIComponent(contactEmail);
  var enrollResp = UrlFetchApp.fetch(enrollUrl, {
    method: "post",
    headers: { Authorization: "Bearer " + token },
    muteHttpExceptions: true
  });
  var enrollCode = enrollResp.getResponseCode();
  if (enrollCode === 204) {
    return { ok: true };
  }
  return { ok: false, error: "Enrollment failed (HTTP " + enrollCode + "): " + enrollResp.getContentText().slice(0, 200) };
}

// ── DEBUG: list a contact's existing tasks ─────────────────────────
// Run this directly from the Apps Script editor (select debugListContactTasks
// in the function dropdown, click Run) to see real task titles/dates HubSpot
// sequences have already created for a contact. Check View > Logs (or
// View > Execution log) afterward to see the output.
//
// This tries the CRM v4 associations endpoint first (contacts -> tasks).
// If your scopes don't include task access, this will throw a clear error
// naming the missing scope — that tells us what to add.
function debugListContactTasks() {
  // EDIT THIS: paste a real contact URL, ID, or leave as a plain ID string
  var contactIdOrUrl = "PASTE_A_CONTACT_URL_OR_ID_HERE";

  var contactId = extractContactId(contactIdOrUrl) || contactIdOrUrl;
  if (!contactId) {
    Logger.log("Couldn't resolve a contact ID from: " + contactIdOrUrl);
    return;
  }

  var assoc;
  try {
    assoc = hsFetch("/crm/v4/objects/contacts/" + contactId + "/associations/tasks", "get");
  } catch (e) {
    Logger.log("Associations call failed: " + e.message);
    Logger.log("This likely means the token is missing a tasks-related scope. The error above should name it.");
    return;
  }

  var results = assoc.results || [];
  if (!results.length) {
    Logger.log("No tasks found associated with contact " + contactId + ".");
    return;
  }

  Logger.log("Found " + results.length + " task(s) for contact " + contactId + ":");
  Logger.log("----------------------------------------------------");

  for (var i = 0; i < results.length; i++) {
    var taskId = results[i].toObjectId || results[i].id;
    if (!taskId) continue;
    try {
      var task = hsFetch("/crm/v3/objects/tasks/" + taskId + "?properties=hs_task_subject,hs_task_body,hs_task_status,hs_task_type,hs_timestamp,hubspot_owner_id", "get");
      var tp = task.properties || {};
      Logger.log("Task ID: " + taskId);
      Logger.log("  Subject: " + (tp.hs_task_subject || "(none)"));
      Logger.log("  Type: " + (tp.hs_task_type || "(none)"));
      Logger.log("  Status: " + (tp.hs_task_status || "(none)"));
      Logger.log("  Due (hs_timestamp): " + (tp.hs_timestamp || "(none)"));
      Logger.log("  Owner ID: " + (tp.hubspot_owner_id || "(none)"));
      Logger.log("----------------------------------------------------");
    } catch (e) {
      Logger.log("Failed to fetch task " + taskId + ": " + e.message);
    }
  }
}

// ── Fathom ────────────────────────────────────────────────────────
// Fetches EVERY page of meetings this API key can see, following
// next_cursor until Fathom stops returning one. Guarded by:
//   - a hard page cap (500 pages ≈ 25,000 meetings) so a malformed
//     response can never loop forever
//   - a wall-clock cutoff at 4.5 minutes, since Apps Script kills any
//     script running past 6 minutes — if we hit the cutoff we stop
//     and return whatever we've gathered so far rather than crashing
// daysBack: only fetch meetings created in the last N days (default 60).
// This is both faster AND avoids rate-limit issues, since onboard calls
// are always recent relative to when this tool runs — no need to page
// through a team's entire multi-year meeting history every time.
function fetchFathomPages(daysBack) {
  daysBack = daysBack || 60;
  var key = getProp("FATHOM_API_KEY");
  var all = [];
  var cursor = null;
  var startTime = new Date().getTime();
  var maxRuntimeMs = 4.5 * 60 * 1000;
  var maxPages = 500;
  var delayMs = 1100; // stay safely under Fathom's 60 calls/minute limit

  var createdAfter = new Date();
  createdAfter.setDate(createdAfter.getDate() - daysBack);
  var createdAfterIso = createdAfter.toISOString();

  for (var p = 0; p < maxPages; p++) {
    if (new Date().getTime() - startTime > maxRuntimeMs) {
      Logger.log("fetchFathomPages: hit time cutoff after " + p + " pages (" + all.length + " meetings) — stopping early.");
      break;
    }

    var url = "https://api.fathom.ai/external/v1/meetings?limit=100"
      + "&created_after=" + encodeURIComponent(createdAfterIso)
      + (cursor ? "&cursor=" + encodeURIComponent(cursor) : "");

    var data = null;
    var retries = 0;
    var maxRetries = 4;

    while (retries <= maxRetries) {
      var resp = UrlFetchApp.fetch(url, {
        method: "get",
        headers: { "X-Api-Key": key },
        muteHttpExceptions: true
      });
      var code = resp.getResponseCode();

      if (code === 429) {
        var retryAfterHeader = resp.getHeaders()["Retry-After"] || resp.getHeaders()["retry-after"];
        var waitMs = retryAfterHeader ? (parseInt(retryAfterHeader, 10) * 1000) : (5000 * (retries + 1));
        Logger.log("Fathom 429 (rate limited) on page " + p + " — waiting " + (waitMs / 1000) + "s before retry " + (retries + 1) + "/" + maxRetries + ".");
        Utilities.sleep(waitMs);
        retries++;
        continue;
      }

      if (code < 200 || code >= 300) {
        throw new Error("Fathom " + code + ": " + resp.getContentText().slice(0, 300));
      }

      data = JSON.parse(resp.getContentText());
      break; // success
    }

    if (!data) {
      throw new Error("Fathom rate limit (429) persisted after " + maxRetries + " retries. Try again in a minute.");
    }

    var items = data.items || [];
    all = all.concat(items);
    cursor = data.next_cursor || null;

    if (!cursor || items.length === 0) break; // no more pages

    if (cursor) Utilities.sleep(delayMs); // pace the next request
  }
  return all;
}

// Searches Fathom meetings page-by-page and stops the INSTANT a match is
// found — unlike fetchFathomPages() (used by the debug tools, which
// deliberately fetches everything so you can see the full picture), this
// is what the real app uses, so it should usually finish in 1-2 API calls
// instead of paging through the whole 60-day window every time.
// Uses Fathom's calendar_invitees[] filter so FATHOM does the matching
// server-side by email — we only get back meetings that actually involved
// this person, regardless of how long ago it happened. This replaced an
// earlier version that fetched a date-windowed page of ALL meetings and
// grepped through them client-side; filtering server-side is both much
// faster and no longer needs an arbitrary "only search the last 60 days"
// cutoff, since the result set is already small and targeted.
//
// If no email is available (name-only), we fall back to a recent-window
// scan since Fathom's filter here works on email, not name.
// Returns { firstDate, firstOnboarder, lastDate, lastOnboarder, totalCount }
// counting ONLY genuine onboarding calls (see isOnboardingMeeting) — demos
// are filtered out entirely — sorted by date so "First Onboard Date" is the
// EARLIEST onboarding call and "Last Onboarding Call Date" is the most
// recent one, with the onboarder taken from the earliest onboarding call.
//
// This distinction matters: an earlier version counted every meeting the
// person appeared in (including demos), which both inflated the call count
// and could mislabel a demo as the "first onboard" date/onboarder.
// ── Fathom call-summary extraction ──────────────────────────────────
// Confirmed against Fathom's published OpenAPI spec (developers.fathom.ai):
// a meeting's AI summary lives at default_summary.markdown_formatted, and
// is ONLY included in the response if the request explicitly passed
// include_summary=true (every Fathom fetch URL in this file sets that).
// Kept defensive (falls back to a couple of alternate shapes) in case the
// live response ever differs slightly from the published spec.
function extractSummary(m) {
  if (!m) return "";
  var s = (m.default_summary && m.default_summary.markdown_formatted) || m.summary || m.ai_summary || "";
  if (s && typeof s === "object") s = s.markdown_formatted || s.markdown || s.text || "";
  return String(s || "").trim().slice(0, 800); // keep prompt small/cheap
}

// Counts distinct EXTERNAL attendees on a meeting, using Fathom's own
// calendar_invitees list and its own is_external flag — Fathom already
// determines internal-vs-external per invitee itself (based on the
// recording account's own domain), so this doesn't need to guess at CRE
// OneSource's own domain or maintain a separate list of team emails.
// Returns null (not 0) when no invitee data is available at all, so the
// caller can distinguish "confirmed exactly one external attendee" from
// "we don't actually know" rather than silently treating both the same.
// Note this reflects who was INVITED via the calendar event, not
// necessarily who actually attended — Fathom doesn't expose a separate
// confirmed-attendance list, so this is the best available signal, not a
// guarantee.
function countExternalAttendees(meeting) {
  var invitees = (meeting && meeting.calendar_invitees) || [];
  if (!Array.isArray(invitees) || !invitees.length) return null;
  var seen = {};
  invitees.forEach(function(inv) {
    if (inv && inv.is_external === true && inv.email) {
      seen[String(inv.email).trim().toLowerCase()] = true;
    }
  });
  return Object.keys(seen).length;
}

// Reads Fathom's Retry-After header off a 429 response (seconds, per HTTP
// spec) and converts it to milliseconds. Falls back to `fallbackMs` if the
// header is missing, malformed, or the response object doesn't expose
// headers for some reason — so a rate-limit backoff always has SOME wait
// time even when the header can't be read, rather than looping instantly.
function parseRetryAfterMs(resp, fallbackMs) {
  try {
    var headers = resp.getHeaders() || {};
    var h = headers["Retry-After"] || headers["retry-after"];
    if (h) {
      var secs = parseInt(h, 10);
      if (!isNaN(secs) && secs >= 0) return secs * 1000;
    }
  } catch (e) {
    // fall through to fallback
  }
  return fallbackMs;
}

function getFathomMeetingRange(personName, personEmail, personCompany, signupDateIso, knownOnboardingDateIso, knownLastOnboardingDateIso, dealCreatedDateIso, knownPastTrainingCount, pastDemoDateIso, suggestedUserTypeFromDealSize, likelyRepNameHint, pastDemoStandIn) {
  // Overall time budget for the WHOLE Fathom resolution (all tiers
  // combined), not a hard execution-kill workaround — confirmed directly
  // that a single execution can legitimately run past 6 minutes and still
  // complete successfully on this account, so this isn't about dodging a
  // ceiling. It's about giving a bounded, predictable wait instead of an
  // open-ended one: confirmed real case, Contact K's search ran over
  // 5 minutes before completing. Once this budget is exceeded, each tier
  // stops gracefully and uses whatever it's found so far, the same way
  // the existing dry-spell/rate-limit stopping points already work,
  // rather than continuing indefinitely.
  var overallSearchStartMs = Date.now();
  var OVERALL_SEARCH_BUDGET_MS = 4 * 60 * 1000; // 4 minutes
  function overallBudgetExceeded() {
    return (Date.now() - overallSearchStartMs) > OVERALL_SEARCH_BUDGET_MS;
  }

  var empty = { firstDate: "", firstOnboarder: "", lastDate: "", lastOnboarder: "", totalCount: 0, salesperson: "",
                demoTitle: "", demoSummary: "", onboardingTitle: "", onboardingSummary: "", suggestedUserType: suggestedUserTypeFromDealSize || "" };
  var key;
  try {
    key = getProp("FATHOM_API_KEY");
  } catch (e) {
    return empty;
  }

  var emailLc = (personEmail || "").trim().toLowerCase();
  var nameLc = (personName || "").toLowerCase();

  // ── Shared paced fetch for every Fathom request in this function ──────
  // All three search tiers below (exact-email, company-domain, and the
  // per-rep fallback) route every single request through this ONE helper,
  // which tracks a single shared "time of last request" across ALL of
  // them — not just within one tier's own loop. This guarantees uniform
  // pacing even at tier hand-offs (the exact-email tier's last request and
  // the domain tier's first request still stay spaced apart), and gives
  // every tier the same real Retry-After handling instead of each tier
  // having its own ad hoc backoff (an earlier version's exact-email tier
  // used a fixed exponential guess instead of honoring Fathom's actual
  // Retry-After header, inconsistent with the other tiers).
  var fathomLastRequestAt = 0;
  function pacedFathomFetch(url, retriesLeft) {
    if (typeof retriesLeft !== "number") retriesLeft = 5;
    var sinceLastMs = fathomLastRequestAt ? (new Date().getTime() - fathomLastRequestAt) : Infinity;
    if (sinceLastMs < FATHOM_PACE_MS_PER_REQUEST) {
      Utilities.sleep(FATHOM_PACE_MS_PER_REQUEST - sinceLastMs);
    }
    fathomLastRequestAt = new Date().getTime();
    var resp = UrlFetchApp.fetch(url, { method: "get", headers: { "X-Api-Key": key }, muteHttpExceptions: true });
    var code = resp.getResponseCode();
    if (code === 429) {
      if (retriesLeft <= 0) {
        traceLog("Fathom", "Rate limited and out of retries for this request — giving up on it", "fail");
        return null;
      }
      var retryMs = parseRetryAfterMs(resp, 5000);
      traceLog("Fathom", "Rate limited — backing off " + (retryMs / 1000) + "s per Fathom's Retry-After before retrying", "warn");
      Utilities.sleep(retryMs);
      fathomLastRequestAt = new Date().getTime(); // the backoff itself satisfies the pacing gap
      return pacedFathomFetch(url, retriesLeft - 1);
    }
    if (code < 200 || code >= 300) return null;
    try {
      return JSON.parse(resp.getContentText());
    } catch (e) {
      return null;
    }
  }

  // ── Known-onboarding-date search window (tightest, highest priority) ──
  // When HubSpot's own Meetings tab already found a candidate training
  // call, we know almost exactly when it happened — even if we don't yet
  // trust WHO HubSpot says hosted it (that's exactly what this Fathom
  // check is for: HubSpot's meeting-owner field can be wrong, e.g. it
  // names whoever the record is filed under rather than who actually ran
  // the call). A tiny window around that known date (covering the case
  // it's a few days off from a reschedule) still lets Fathom's
  // recorded_by fully verify — and correct — the onboarder, but turns a
  // 90-day, all-reps scan into a handful of near-instant single-page
  // requests. This does NOT skip verification the way the old "trust
  // HubSpot and skip Fathom" shortcut did — it only shrinks the search
  // space using a fact we already know, and Fathom's own recorded_by
  // still has final say over both the onboarder name and whether the call
  // even counts (a future, unrecorded meeting simply won't show up here).
  var windowCreatedAfterIso = null;
  var windowCreatedBeforeIso = null;
  if (knownOnboardingDateIso) {
    var knownDate = new Date(knownOnboardingDateIso + "T00:00:00Z");
    if (!isNaN(knownDate.getTime())) {
      var kStart = new Date(knownDate.getTime());
      kStart.setDate(kStart.getDate() - 5);

      // Upper bound spans to the LATEST known training-call date (+5),
      // not just +5 days past the EARLIEST one. A fixed +5-around-
      // earliest window silently missed later real calls for anyone
      // whose onboarding stretched out longer than 5 days — confirmed
      // as a real bug: a contact with 4 genuine recorded calls spanning
      // two full weeks had 3 of them fall outside a window anchored only
      // on the first. Falling back to the earliest date +5 when no later
      // date is known keeps the original tight behavior for the common
      // single-call case.
      var kEnd;
      if (knownLastOnboardingDateIso) {
        var knownLastDate = new Date(knownLastOnboardingDateIso + "T00:00:00Z");
        kEnd = !isNaN(knownLastDate.getTime()) ? new Date(knownLastDate.getTime()) : new Date(knownDate.getTime());
      } else {
        kEnd = new Date(knownDate.getTime());
      }
      kEnd.setDate(kEnd.getDate() + 5);

      windowCreatedAfterIso = kStart.toISOString();
      windowCreatedBeforeIso = kEnd.toISOString();
      traceLog("Fathom Window", "Using a known-onboarding-date search window: " + isoDate(kStart) + " to " + isoDate(kEnd) + " (HubSpot's candidate training call date" + (knownLastOnboardingDateIso && knownLastOnboardingDateIso !== knownOnboardingDateIso ? "s: " + knownOnboardingDateIso + " through " + knownLastOnboardingDateIso : ": " + knownOnboardingDateIso) + ") — verifying via Fathom's recorded_by rather than trusting HubSpot's meeting owner", "ok");
    }
  }

  // ── Signup-anchored search window ─────────────────────────────────
  // When the AppData tab gives us a real sign-up date, search a TIGHT
  // window anchored to it — a demo happens shortly BEFORE sign-up, and
  // onboarding calls happen shortly AFTER — instead of a blind window
  // measured back from today. This is both faster (far fewer pages to
  // page through) AND more correct: a contact who signed up months ago
  // could fall entirely outside a fixed "today minus N days" window,
  // meaning Fathom would never find their calls no matter how long the
  // scan ran. Falls back to the old today-anchored window when no
  // sign-up date is available (e.g. AppData didn't have a match). Only
  // used when the known-onboarding-date window above wasn't available.
  if (!windowCreatedAfterIso && signupDateIso) {
    var signup = new Date(signupDateIso + "T00:00:00Z");
    if (!isNaN(signup.getTime())) {
      var winStart = new Date(signup.getTime());
      winStart.setDate(winStart.getDate() - 30); // demo can precede signup by ~a month
      var winEnd = new Date(signup.getTime());
      winEnd.setDate(winEnd.getDate() + 60); // onboarding calls happen within ~2 months of signup
      windowCreatedAfterIso = winStart.toISOString();
      windowCreatedBeforeIso = winEnd.toISOString();
      traceLog("Fathom Window", "Using a sign-up-anchored search window: " + isoDate(winStart) + " to " + isoDate(winEnd) + " (signup date: " + signupDateIso + ")", "ok");
    }
  }
  var windowIsQuickGuess = false;
  if (!windowCreatedAfterIso && pastDemoDateIso) {
    // A past demo/generic call is standing in for onboarding because
    // HubSpot's Meetings tab has no dedicated training call for this
    // contact — but that doesn't mean nothing real exists in Fathom.
    // HubSpot's own record can be genuinely incomplete: confirmed real
    // case — a team's actual onboarding sessions were logged under a
    // THIRD person's calendar (a team coordinator helping run setup),
    // never appearing on either real contact's own Meetings tab at all,
    // even though multiple real, dated recordings existed in Fathom the
    // whole time, matched correctly by attendee once actually searched
    // for. This window searches from just before the known demo through
    // today — wide enough to catch onboarding activity that happened any
    // time after the demo, however long it took to actually get started.
    // Marked as a quick guess (not a confirmed date) for the same reason
    // the "just happened" window below is: if nothing turns up, Tier 3's
    // fallback below should use its own proper wide scan instead of
    // being wrongly restricted to this same guessed range.
    var demoDate = new Date(pastDemoDateIso + "T00:00:00Z");
    if (!isNaN(demoDate.getTime())) {
      var pdStart = new Date(demoDate.getTime());
      pdStart.setDate(pdStart.getDate() - 5);
      var pdEnd = new Date();
      pdEnd.setDate(pdEnd.getDate() + 2);
      windowCreatedAfterIso = pdStart.toISOString();
      windowCreatedBeforeIso = pdEnd.toISOString();
      windowIsQuickGuess = true;
      traceLog("Fathom Window", "No dedicated training call on HubSpot's Meetings tab — searching from just before the past-call stand-in (" + pastDemoDateIso + ") through today (" + isoDate(pdStart) + " to " + isoDate(pdEnd) + "), in case real onboarding activity exists in Fathom that HubSpot's own record doesn't know about", "ok");
    }
  }
  if (!windowCreatedAfterIso) {
    // No HubSpot candidate date, no AppData sign-up date — the most
    // likely real-world reason for BOTH of those being blank at once is
    // exactly the case Lupe is actually built for: a brand-new contact,
    // pulled right after their first-ever call, before HubSpot's Meetings
    // tab has synced anything and before AppData has a signup date on
    // file yet. Rather than leave Tier 1/2 completely unbounded (scanning
    // this contact's ENTIRE Fathom history), try a tight "just happened"
    // window first — today, plus a day of buffer on each side for
    // timezone edges.
    //
    // windowIsQuickGuess marks this as an UNCONFIRMED guess, not a real
    // hint — this matters because Tier 3 below normally REUSES whatever
    // window Tier 1/2 used rather than falling back to its own scan.
    // Without this flag, a wrong guess here (testing an older contact
    // with no other date hints available) would incorrectly restrict
    // Tier 3's fallback scan to just this same narrow range too, causing
    // it to find NOTHING even though its own proper 90-day scan would
    // have found the call just fine. The flag tells Tier 3 to ignore this
    // guess and fall back to its own independent range instead, exactly
    // as if no window had been set at all.
    windowIsQuickGuess = true;
    var todayGuess = new Date();
    var todayStart = new Date(todayGuess.getTime());
    todayStart.setDate(todayStart.getDate() - 4);
    var todayEnd = new Date(todayGuess.getTime());
    todayEnd.setDate(todayEnd.getDate() + 1);
    windowCreatedAfterIso = todayStart.toISOString();
    windowCreatedBeforeIso = todayEnd.toISOString();
    traceLog("Fathom Window", "No HubSpot candidate date or sign-up date available — trying a quick just-happened window first: " + isoDate(todayStart) + " to " + isoDate(todayEnd) + " (falls through to the full " + FATHOM_SCAN_DAYS_BACK + "-day scan below if this finds nothing)", "info");
  }

  // Split the (already-resolved, cleaned) full name into first/last for
  // attendee matching. Using the resolved name avoids the dirty raw
  // firstname/lastname fields some records have.
  var nameParts = (personName || "").trim().split(/\s+/).filter(function(w){ return w; });
  var firstName = nameParts.length ? nameParts[0] : "";
  var lastName  = nameParts.length > 1 ? nameParts[nameParts.length - 1] : "";
  var coTokens  = companyTokens(personCompany);

  // A meeting belongs to this contact if the attendee matcher confirms it
  // (email present, first+last present, or first-name + company present).
  var involvesContact = function(m) {
    return contactMatchesMeeting(m, firstName, lastName, emailLc, coTokens);
  };

  var resolveHost = function(m, roleLabel) {
    var host = m.recorded_by && m.recorded_by.name ? m.recorded_by.name : "";
    // Only accept the host if they're a valid salesperson/onboarder; a
    // non-rep host (e.g. a cold-caller who sat in) is rejected -> "".
    var validated = validRep(host);
    if (host && !validated) {
      traceLog("Fathom→" + (roleLabel || "Host"), 'Recorded-by "' + host + '" is not on the valid-rep roster — rejected', "warn");
    }
    return validated;
  };
  // Given a list of candidate meetings ALREADY SORTED oldest-first, returns
  // the first one whose Fathom host (recorded_by) actually validates
  // against the roster — instead of only ever checking index 0. This
  // mirrors firstValidRepMeeting() above (the HubSpot-side version) but
  // works against Fathom meeting objects via resolveHost(), which already
  // does the roster-validation + rejection logging.
  var firstValidHostMeeting = function(sortedMeetings, roleLabel) {
    for (var i = 0; i < sortedMeetings.length; i++) {
      var host = resolveHost(sortedMeetings[i], roleLabel);
      if (host) return { meeting: sortedMeetings[i], name: host, index: i };
    }
    return null;
  };
  var resolveDate = function(m) {
    var s = m.scheduled_start_time || m.recording_start_time || m.created_at || "";
    return s ? isoDate(new Date(s)) : "";
  };
  var resolveTimestamp = function(m) {
    var s = m.scheduled_start_time || m.recording_start_time || m.created_at || "";
    var t = s ? new Date(s).getTime() : NaN;
    return isNaN(t) ? 0 : t;
  };
  // Classifies the contact's meetings into onboarding calls and demos.
  // Onboarding info (first/last DATE and total COUNT) always reflects the
  // actual earliest/latest/total onboarding calls found, regardless of who
  // hosted them — that's a factual record of what happened. The ASSIGNED
  // onboarder/salesperson, and the title/summary shown as evidence for
  // that assignment, come from the first candidate (in chronological
  // order) whose host actually validates against the roster — via
  // firstValidHostMeeting() — rather than only ever checking the single
  // earliest meeting and giving up if THAT one happens to be hosted by a
  // non-roster person (e.g. an inactive rep sitting in before the real
  // onboarder joined). Returns empty (still carrying any salesperson we
  // found) if there are no onboarding calls, so a demo is never mistaken
  // for an onboard.
  var summarize = function(items, sourceLabel) {
    var label = sourceLabel || "Fathom";
    var deduped = dedupeMeetings(items, resolveTimestamp);
    if (deduped.length !== items.length) {
      traceLog("Fathom Dedupe", (items.length - deduped.length) + ' duplicate bot recording(s) collapsed (same title + same call time, e.g. two reps both recording the same Zoom)', "warn");
    }

    var demos = deduped.filter(isDemoMeeting);
    var salesperson = "";
    var demoTitle = "", demoSummary = "";
    if (demos.length) {
      demos.sort(function(a, b) { return resolveTimestamp(a) - resolveTimestamp(b); });
      var demoPick = firstValidHostMeeting(demos, "Salesperson");
      var demoEvidence = demoPick ? demoPick.meeting : demos[0];
      demoTitle = demoEvidence.meeting_title || demoEvidence.title || "";
      demoSummary = extractSummary(demoEvidence);
      traceLog("Fathom Classify", '[' + label + '] Earliest demo: "' + (demoTitle || "(untitled)") + '"' + (demoSummary ? " (summary retrieved)" : " (no summary in response)"), "info");
      salesperson = demoPick ? demoPick.name : "";
      if (salesperson) traceLog("Fathom→Salesperson", 'Demo host "' + salesperson + '" — valid rep ✓ SELECTED (candidate ' + (demoPick.index + 1) + ' of ' + demos.length + ')', "ok");
    }

    var onboardings = deduped.filter(function(m) { return isOnboardingMeeting(m, firstName); });
    if (!onboardings.length) {
      // No dedicated training-titled call exists in this window — but if
      // HubSpot's own side already decided (via getMeetingRepsFromHubSpot)
      // that a demo-titled call is standing in as the real onboarding
      // session, mirror that decision on the Fathom side too instead of
      // reporting "no onboarding call found" and falling through to the
      // full multi-tier fallback scan. That scan would burn through rate
      // limits chasing a classification ("onboarding") this call's title
      // can never satisfy — confirmed real case: Contact C's only
      // real onboarding session is titled "CRE OneSource Meeting w/ Contact C" (a demo-style title), so isOnboardingMeeting() will
      // never match it no matter how long the scan runs.
      if (pastDemoStandIn && demoPick) {
        var standInTitle = demoPick.meeting.meeting_title || demoPick.meeting.title || "";
        traceLog("Fathom Classify", '[' + label + '] No dedicated training call exists — treating demo-titled call "' + (standInTitle || "(untitled)") + '" as the real onboarding session instead (matches HubSpot\'s own past-demo stand-in decision)', "ok");
        traceLog("Fathom→Onboarder", 'Host "' + demoPick.name + '" — valid rep ✓ SELECTED as onboarder from demo stand-in', "ok");
        var standInSalesperson = salesperson;
        if (demoPick.meeting) {
          traceLog("Fathom→Salesperson", 'Retracting salesperson "' + salesperson + '" — it came from this SAME meeting, now being used as the onboarding stand-in instead', "warn");
          standInSalesperson = "";
        }
        return {
          firstDate: resolveDate(demoPick.meeting),
          firstOnboarder: demoPick.name,
          lastDate: resolveDate(demoPick.meeting),
          lastOnboarder: demoPick.name,
          totalCount: 1,
          salesperson: standInSalesperson,
          demoTitle: "",
          demoSummary: "",
          onboardingTitle: standInTitle,
          onboardingSummary: extractSummary(demoPick.meeting),
          suggestedUserType: suggestedUserTypeFromDealSize || ""
        };
      }
      traceLog("Fathom Classify", '[' + label + '] No onboarding-titled meetings among ' + deduped.length + ' candidate(s)', "warn");
      return { firstDate: "", firstOnboarder: "", lastDate: "", lastOnboarder: "", totalCount: 0, salesperson: salesperson,
               demoTitle: demoTitle, demoSummary: demoSummary, onboardingTitle: "", onboardingSummary: "", suggestedUserType: suggestedUserTypeFromDealSize || "" };
    }
    onboardings.sort(function(a, b) { return resolveTimestamp(a) - resolveTimestamp(b); });
    var first = onboardings[0];
    var last = onboardings[onboardings.length - 1];
    var onboardingPick = firstValidHostMeeting(onboardings, "Onboarder");
    var onboardingEvidence = onboardingPick ? onboardingPick.meeting : first;
    var onboardingTitle = onboardingEvidence.meeting_title || onboardingEvidence.title || "";
    var onboardingSummary = extractSummary(onboardingEvidence);
    traceLog("Fathom Classify", '[' + label + '] ' + onboardings.length + ' onboarding call(s) found; earliest: "' + (first.meeting_title || first.title || "(untitled)") + '"; using "' + (onboardingTitle || "(untitled)") + '" as onboarder evidence' + (onboardingSummary ? " (summary retrieved)" : " (no summary in response)"), "info");
    var firstOnboarder = onboardingPick ? onboardingPick.name : "";
    if (firstOnboarder) traceLog("Fathom→Onboarder", 'Host "' + firstOnboarder + '" — valid rep ✓ SELECTED (candidate ' + (onboardingPick.index + 1) + ' of ' + onboardings.length + ')', "ok");
    // Secondary cross-reference against the "Deal Size" deal field (the
    // PRIMARY signal, applied immediately in the fast path — see
    // getContactDeal). This attendee-based check runs regardless, but its
    // role now depends on whether Deal Size was actually set:
    //   - Deal Size present, and this AGREES → just confirms it.
    //   - Deal Size present, and this DISAGREES → flagged as a real
    //     discrepancy worth a manual look, but does NOT override Deal
    //     Size, since that reflects what your sales team actually
    //     recorded — an attendee list only reflects who was invited to
    //     ONE call.
    //   - Deal Size not set at all → this becomes the fallback
    //     suggestion, exactly as it worked before this cross-reference
    //     existed.
    var externalAttendeeCount = countExternalAttendees(onboardingEvidence);
    var attendeeBasedType = (externalAttendeeCount === null) ? "" : (externalAttendeeCount > 1 ? "Team" : "Individual");
    var suggestedUserType = suggestedUserTypeFromDealSize || attendeeBasedType;
    if (externalAttendeeCount !== null) {
      if (suggestedUserTypeFromDealSize && attendeeBasedType && suggestedUserTypeFromDealSize !== attendeeBasedType) {
        traceLog("Fathom→User Type", "⚠ Mismatch: \"Deal Size\" on the deal suggests \"" + suggestedUserTypeFromDealSize + "\", but " + externalAttendeeCount + " external attendee(s) on the actual onboarding call suggests \"" + attendeeBasedType + "\" — keeping Deal Size (what your sales team actually recorded), but worth a manual look", "warn");
      } else if (suggestedUserTypeFromDealSize) {
        traceLog("Fathom→User Type", externalAttendeeCount + " external attendee(s) on the onboarding call — agrees with the \"" + suggestedUserTypeFromDealSize + "\" already suggested from \"Deal Size\"", "ok");
      } else {
        traceLog("Fathom→User Type", externalAttendeeCount + " external attendee(s) on the onboarding call — no \"Deal Size\" was set on the deal, so suggesting \"" + attendeeBasedType + "\" from attendees instead (editable)", "info");
      }
      // Zero-cost attendance cross-check: reuses the summary text already
      // extracted above (extractSummary) — no new API call, no extra
      // pacing delay, since this data is already sitting in memory for
      // display purposes anyway. Only relevant when Team is the
      // attendee-based read specifically — an Individual call has nobody
      // else to be confused with, so there's nothing to disambiguate. A
      // real transcript speaker-list check would be a stronger, more
      // certain signal than a summary mention, but costs a meaningfully
      // heavier, separately-paced API call per Fathom's own guidance
      // ("transcripts are large") — this catches the same concern for
      // free by checking data already in hand, at the cost of being a
      // softer heuristic (a summary can mention someone as the subject of
      // discussion without them having personally spoken, and can omit a
      // quiet attendee who said little).
      if (attendeeBasedType === "Team" && firstName && onboardingSummary) {
        var nameInSummary = onboardingSummary.toLowerCase().indexOf(firstName.toLowerCase()) >= 0;
        if (!nameInSummary) {
          traceLog("Fathom→User Type", "⚠ \"" + firstName + "\" isn't mentioned by name in this call's summary, despite " + externalAttendeeCount + " external attendees being invited — worth confirming they were actually on this specific call", "warn");
        }
      }
    }
    return {
      firstDate: resolveDate(first),
      firstOnboarder: firstOnboarder,
      lastDate: resolveDate(last),
      lastOnboarder: resolveHost(last, "Onboarder"),
      totalCount: onboardings.length,
      salesperson: salesperson,
      demoTitle: demoTitle,
      demoSummary: demoSummary,
      onboardingTitle: onboardingTitle,
      onboardingSummary: onboardingSummary,
      suggestedUserType: suggestedUserType
    };
  };

  // ── Tier 1: exact-email invitee filter, but VERIFY every result ourselves ──
  // Fathom's calendar_invitees[] filter is used as a speed optimization,
  // but we don't blindly trust it — if a returned meeting doesn't actually
  // contain this email anywhere in its own data, we throw it out. This
  // guards against the filter being ignored/misapplied server-side, which
  // was observed returning an unrelated person's most recent meeting
  // instead of an empty result.
  if (emailLc) {
    try {
      var allItems = [];
      var cursor = null;
      var maxPages = 20;

      // Grace-period early exit: once a usable onboarding call has been
      // verified, keep paging for a FEW more pages (to still catch a
      // genuine nearby second call — the reason full pagination was
      // added in the first place), then stop rather than paging
      // unconditionally to maxPages. An earlier version removed early
      // exit entirely on the theory that this tier's result set is
      // always small (server-side filtered to one exact email) — but
      // that assumption broke in practice: one contact's email matched
      // 100+ meetings within just a 10-day window, and page 2 alone hit
      // 5 consecutive Fathom rate-limit retries before giving up,
      // costing ~36 seconds for zero additional benefit (nothing new was
      // ever found on that page). A bounded grace period gets most of
      // the completeness benefit without the unbounded cost.
      var TIER1_GRACE_PAGES = 2;
      var TIER1_GRACE_PAGES_WHILE_CHASING_TARGET = 6;
      var lastNewMatchPage = -1;
      var lastOnboardingCount = 0;
      for (var p = 0; p < maxPages; p++) {
        if (overallBudgetExceeded()) {
          traceLog("Fathom Invitee-Filter", "Hit the overall " + (OVERALL_SEARCH_BUDGET_MS / 1000) + "s search budget — stopping here with whatever's been found so far instead of continuing indefinitely", "warn");
          break;
        }
        var url = "https://api.fathom.ai/external/v1/meetings?limit=100&include_summary=true&include_crm_matches=true"
          + "&calendar_invitees[]=" + encodeURIComponent(emailLc)
          + (windowCreatedAfterIso ? "&created_after=" + encodeURIComponent(windowCreatedAfterIso) : "")
          + (windowCreatedBeforeIso ? "&created_before=" + encodeURIComponent(windowCreatedBeforeIso) : "")
          + (cursor ? "&cursor=" + encodeURIComponent(cursor) : "");

        var data = pacedFathomFetch(url);
        if (!data) break;

        var items = data.items || [];
        allItems = allItems.concat(items);

        // Reset the grace countdown every time the onboarding-classified
        // count actually GROWS, not just once from the first match — a
        // fixed "N pages after the FIRST match" grace period assumes a
        // contact's real calls cluster together in the pagination, which
        // isn't true when their email is invited to many unrelated
        // meetings (calendar_invitees[] matches ANY meeting they're on,
        // not just onboarding-related ones): confirmed real case — Contact L's 3 genuine calls were scattered across 70+ invitee-matched
        // meetings, with the first one found on an early page but the
        // other two sitting several pages further out. Stopping 2 pages
        // after only the FIRST discovery missed both of them. Extending
        // the grace window every time a NEW onboarding call appears keeps
        // chasing genuinely active pages, while still cutting off once a
        // real dry spell of consecutive empty pages passes.
        var verifiedSoFar = allItems.filter(involvesContact);
        var currentOnboardingCount = verifiedSoFar.length ? summarize(verifiedSoFar, "Fathom invitee-filter (grace check)").totalCount : 0;
        if (currentOnboardingCount > lastOnboardingCount) {
          lastOnboardingCount = currentOnboardingCount;
          lastNewMatchPage = p;
        }

        cursor = data.next_cursor || null;
        if (!cursor || items.length === 0) break;

        // Confident, IMMEDIATE stop — no need to wait out any grace pages
        // at all — once we've found as many onboarding calls as HubSpot's
        // own Meetings tab already told us to expect. This is the actual
        // fix for the tension a flat grace-page count can't resolve: a
        // fixed number is either too aggressive for a simple one-call
        // contact (wastes time chasing pages that were never going to
        // reveal anything, as happened with Mark Elrod) or not generous
        // enough for a contact with several genuinely scattered real
        // calls (missed 2 of Contact L's 3). Knowing the real target
        // count up front removes the guessing entirely for any contact
        // where HubSpot's own data already tells us how many to look for.
        if (knownPastTrainingCount > 0 && currentOnboardingCount >= knownPastTrainingCount) {
          traceLog("Fathom Invitee-Filter", "Found " + currentOnboardingCount + " of " + knownPastTrainingCount + " known training call(s) from HubSpot — stopping confidently, nothing more to look for", "ok");
          break;
        }
        // While we KNOW there should be more (HubSpot told us N training
        // calls exist, but we haven't found N yet), give the search a
        // genuinely generous — but NOT unlimited — extra allowance before
        // giving up, rather than either the plain short dry-spell cutoff
        // OR no cutoff at all. Confirmed as a real, two-sided problem:
        // the plain short cutoff alone missed 2 of Contact L's 3 real
        // calls (scattered deep in a noisy invitee list). But fully
        // suppressing the cutoff assumed HubSpot's own count is always
        // accurate — it isn't: HubSpot doesn't have Fathom's meeting_type
        // tag and can't distinguish a real onboarding call from a
        // similarly-titled non-onboarding one, so its count can OVERcount
        // too. Confirmed real case: Contact H's HubSpot count said 2
        // training calls, but Fathom's more accurate meeting_type
        // classification could only ever confirm 1 real one — an
        // unreachable target, which drove the search all the way to a
        // genuine rate-limit wall (5 retries, ~31s) instead of giving up
        // gracefully. A generous-but-bounded tolerance gets the benefit
        // for genuinely reachable targets without that unbounded risk
        // when the target itself turns out to be wrong.
        var stillChasingKnownTarget = knownPastTrainingCount > 0 && currentOnboardingCount < knownPastTrainingCount;
        var effectiveGracePages = (stillChasingKnownTarget || windowIsQuickGuess) ? TIER1_GRACE_PAGES_WHILE_CHASING_TARGET : TIER1_GRACE_PAGES;
        if (lastNewMatchPage !== -1 && (p - lastNewMatchPage) >= effectiveGracePages) {
          traceLog("Fathom Invitee-Filter", "Found " + lastOnboardingCount + (knownPastTrainingCount > 0 ? " of " + knownPastTrainingCount + " known" : "") + " onboarding call(s) so far, none NEW in the last " + effectiveGracePages + " page(s)" + (stillChasingKnownTarget ? " (extended tolerance, still short of HubSpot's count — likely an overcount on HubSpot's side, e.g. a non-onboarding call sharing a similar title)" : windowIsQuickGuess ? " (extended tolerance — searching an uncertain/guessed window, where HubSpot's own knowledge is already known to be incomplete)" : "") + " — stopping instead of paging all the way to the cap", "ok");
          break;
        }
      }

      // VERIFY with the attendee matcher, not just an email grep — so a
      // meeting the invitee filter returned still counts even if the email
      // shows up only obliquely, and we stay consistent with the fallback.
      var verified = allItems.filter(involvesContact);
      traceLog("Fathom Invitee-Filter", allItems.length + ' meeting(s) returned by calendar_invitees[] filter; ' + verified.length + ' verified as this contact', verified.length ? "info" : "warn");

      if (verified.length) {
        var summary = summarize(verified, "Fathom invitee-filter");
        // Only return here if we actually found an onboarding call. If the
        // verified meetings were all demos (summary is empty), fall through
        // to the company-domain filter in case an onboarding call was
        // missed by the (unreliable) calendar_invitees filter.
        if (summary.totalCount > 0) {
          traceLog("Fathom", "Resolved via invitee-filter path — remaining tiers not needed", "ok");
          return summary;
        }
        traceLog("Fathom", "Invitee-filter found meetings but no onboarding call among them — trying company-domain filter next", "warn");
      }
      // If the filter returned nothing verifiable, don't trust it for this
      // lookup — fall through to the tiers below.
    } catch (e) {
      traceLog("Fathom Invitee-Filter", "Error: " + e.message + " — falling through to company-domain filter", "fail");
    }
  }

  // ── Tier 2: company-domain filter ──────────────────────────────────
  // Catches exactly what the exact-email filter misses: someone who
  // attended a call but wasn't a FORMAL calendar invitee themselves — e.g.
  // Elia's onboarding call was booked under a colleague's invite, so her
  // own email never appeared in calendar_invitees[] no matter how the
  // scan was tuned. Fathom infers ONE associated company per meeting from
  // its invitees' email domains, and calendar_invitees_domains[] filters
  // on that — so a meeting can match here even when this exact person's
  // email was never on the invite. Still verified afterward via the same
  // attendee matcher (a domain match alone only proves SOME meeting
  // involved that company, not that this specific person was on it).
  // Skipped for personal email providers (gmail.com etc.), which aren't a
  // meaningful company scope to filter on.
  var emailDomain = emailLc.indexOf("@") >= 0 ? emailLc.split("@")[1] : "";
  if (emailDomain && !GENERIC_EMAIL_DOMAINS[emailDomain]) {
    try {
      var domainItems = [];
      var domainCursor = null;
      var domainMaxPages = 10;

      // Same grace-period reasoning as Tier 1: page a FEW more times
      // after finding a usable onboarding call (to still catch a genuine
      // nearby second call), then stop — rather than paging
      // unconditionally to domainMaxPages, which proved costly for a
      // contact whose invitee-filter match alone already spanned
      // multiple pages and hit real rate limits.
      var TIER2_GRACE_PAGES = 2;
      var TIER2_GRACE_PAGES_WHILE_CHASING_TARGET = 6;
      var domainLastNewMatchPage = -1;
      var domainLastOnboardingCount = 0;
      for (var dp = 0; dp < domainMaxPages; dp++) {
        if (overallBudgetExceeded()) {
          traceLog("Fathom Domain-Filter", "Hit the overall " + (OVERALL_SEARCH_BUDGET_MS / 1000) + "s search budget — stopping here with whatever's been found so far instead of continuing indefinitely", "warn");
          break;
        }
        var domainUrl = "https://api.fathom.ai/external/v1/meetings?limit=100&include_summary=true&include_crm_matches=true"
          + "&calendar_invitees_domains[]=" + encodeURIComponent(emailDomain)
          + (windowCreatedAfterIso ? "&created_after=" + encodeURIComponent(windowCreatedAfterIso) : "")
          + (windowCreatedBeforeIso ? "&created_before=" + encodeURIComponent(windowCreatedBeforeIso) : "")
          + (domainCursor ? "&cursor=" + encodeURIComponent(domainCursor) : "");

        var domainData = pacedFathomFetch(domainUrl);
        if (!domainData) break;

        var domainPageItems = domainData.items || [];
        domainItems = domainItems.concat(domainPageItems);

        // Same fix as Tier 1: reset the grace countdown every time the
        // onboarding-classified count actually grows, not just once from
        // the first match — see the comment on TIER1_GRACE_PAGES above.
        var domainVerifiedSoFar = domainItems.filter(involvesContact);
        var domainCurrentOnboardingCount = domainVerifiedSoFar.length ? summarize(domainVerifiedSoFar, "Fathom company-domain filter (grace check)").totalCount : 0;
        if (domainCurrentOnboardingCount > domainLastOnboardingCount) {
          domainLastOnboardingCount = domainCurrentOnboardingCount;
          domainLastNewMatchPage = dp;
        }

        domainCursor = domainData.next_cursor || null;
        if (!domainCursor || domainPageItems.length === 0) break;

        if (knownPastTrainingCount > 0 && domainCurrentOnboardingCount >= knownPastTrainingCount) {
          traceLog("Fathom Domain-Filter", "Found " + domainCurrentOnboardingCount + " of " + knownPastTrainingCount + " known training call(s) from HubSpot — stopping confidently, nothing more to look for", "ok");
          break;
        }
        var domainStillChasingKnownTarget = knownPastTrainingCount > 0 && domainCurrentOnboardingCount < knownPastTrainingCount;
        var domainEffectiveGracePages = (domainStillChasingKnownTarget || windowIsQuickGuess) ? TIER2_GRACE_PAGES_WHILE_CHASING_TARGET : TIER2_GRACE_PAGES;
        if (domainLastNewMatchPage !== -1 && (dp - domainLastNewMatchPage) >= domainEffectiveGracePages) {
          traceLog("Fathom Domain-Filter", "Found " + domainLastOnboardingCount + (knownPastTrainingCount > 0 ? " of " + knownPastTrainingCount + " known" : "") + " onboarding call(s) so far, none NEW in the last " + domainEffectiveGracePages + " page(s) — stopping instead of paging all the way to the cap", "ok");
          break;
        }
      }

      var domainVerified = domainItems.filter(involvesContact);
      traceLog("Fathom Domain-Filter", domainItems.length + ' meeting(s) returned by calendar_invitees_domains[]=' + emailDomain + '; ' + domainVerified.length + ' verified as this contact', domainVerified.length ? "info" : "warn");

      if (domainVerified.length) {
        var domainSummary = summarize(domainVerified, "Fathom company-domain filter");
        if (domainSummary.totalCount > 0) {
          traceLog("Fathom", "Resolved via company-domain filter — full per-rep scan not needed", "ok");
          return domainSummary;
        }
        traceLog("Fathom", "Company-domain filter found meetings but no onboarding call among them — trying full per-rep scan", "warn");
      }
    } catch (e) {
      traceLog("Fathom Domain-Filter", "Error: " + e.message + " — falling through to full per-rep scan", "fail");
    }
  } else if (emailDomain) {
    traceLog("Fathom Domain-Filter", 'Skipped — "' + emailDomain + '" is a personal email provider, not a meaningful company scope', "info");
  }

  // ── Tier 3 (fallback): scan a recent window and match by ATTENDEE, not the invitee filter ──
  // This is the path that catches people the calendar_invitees[] filter
  // misses — e.g. someone who attended but wasn't a formal invitee, or who
  // shared one booking with a colleague (title "... Liz & Anna, Denny Elwell
  // Company"). It fetches every meeting in the window and applies the same
  // attendee matcher (email / first+last / first-name+company), then
  // summarize() keeps only onboarding-titled ones.
  if (!nameLc && !emailLc && !coTokens.length) {
    traceLog("Fathom Fallback", "No name/email/company available to match on — skipping fallback scan", "fail");
    return empty;
  }

  var daysBack = FATHOM_SCAN_DAYS_BACK; // configurable at top of file (90 = testing, 30 = production)
  var fbStart = new Date().getTime();
  var fbMaxRuntimeMs = 4.5 * 60 * 1000;
  var fbMatches = [];
  var fbPagesScanned = 0;
  var fbMeetingsScanned = 0;

  // Prefer the signup-anchored or known-date window (tight, correct even
  // for older contacts) computed above; fall back to Tier 3's own
  // independent scan when NO window was available OR when the only
  // window available was just the speculative today-guess above (which,
  // by the time Tier 3 runs, has already failed to find anything in
  // Tiers 1/2 — reusing it here would wrongly restrict this scan to the
  // same narrow range instead of actually searching further back).
  var createdAfterIso, createdBeforeIso;
  if (windowCreatedAfterIso && !windowIsQuickGuess) {
    createdAfterIso = windowCreatedAfterIso;
    createdBeforeIso = windowCreatedBeforeIso;
  } else {
    var fbCreatedAfter = new Date();
    fbCreatedAfter.setDate(fbCreatedAfter.getDate() - daysBack);
    createdAfterIso = fbCreatedAfter.toISOString();
    createdBeforeIso = null;
  }

  // ── Sequential per-rep scan ──────────────────────────────────────────
  // An earlier version of this scan fired multiple reps' requests
  // CONCURRENTLY per round via UrlFetchApp.fetchAll() (first at full
  // concurrency, then capped to 3-at-a-time after that tripped Fathom's
  // limit) — but live testing showed even 3 concurrent requests still got
  // rate-limited on nearly every single round, each costing a real
  // Retry-After backoff. That means Fathom's actual enforcement doesn't
  // tolerate multi-request bursts the way a simple "60/minute budget"
  // would suggest — concurrency itself is the problem, not just sustained
  // rate. So this scan is now fully SEQUENTIAL: one request at a time,
  // paced by FATHOM_PACE_MS_PER_REQUEST between each — the same pattern
  // fetchFathomPages() already uses elsewhere in this file, which has not
  // shown this rate-limiting problem. Reps are scanned one at a time, in
  // roster order; a rep's pages are exhausted before moving to the next.
  // The early-exit checks run after EVERY individual request (not once
  // per "round" of several reps), so a fast match still resolves quickly
  // despite the lack of concurrency.
  var repEmails = VALID_REPS.map(function(name) { return REP_EMAILS[name]; }).filter(function(e) { return e; });

  // Check a known likely candidate FIRST, ahead of the rest of the
  // roster's normal order — HubSpot's own fast-path guess (the deal
  // owner, or whoever's named in a meeting title even if that specific
  // meeting doesn't qualify as onboarding evidence) is often exactly who
  // Fathom needs to check to find the real recording, and there's no
  // reason to burn limited scan time on unrelated reps first when this
  // hint is available. Falls through to the normal roster order
  // unchanged when no hint is given, or the hint doesn't match anyone.
  if (likelyRepNameHint) {
    var hintEmail = REP_EMAILS[likelyRepNameHint];
    if (hintEmail) {
      var hintIdx = repEmails.indexOf(hintEmail);
      if (hintIdx > 0) {
        repEmails.splice(hintIdx, 1);
        repEmails.unshift(hintEmail);
      }
    }
  }

  traceLog("Fathom Fallback", "Starting SEQUENTIAL fallback scan (" + (windowCreatedAfterIso && !windowIsQuickGuess ? "signup-anchored window" : daysBack + "d window from today") + ", " + repEmails.length + " reps scanned one at a time, paced " + FATHOM_PACE_MS_PER_REQUEST + "ms apart) — matching by attendee (email / first+last / first-name+company: " + (coTokens.join(", ") || "none") + ")", "info");

  var fbMaxRequests = 90; // hard cap on total sequential requests this scan will make — generous headroom for a signup-anchored (or 90-day) window across 7 reps
  var requestsMade = 0;
  // Tracks the request count at which an onboarding call was first
  // confirmed among fbMatches, so we can stop waiting on a demo that may
  // not exist at all (or falls outside this scan's reach — e.g. a demo
  // from a year prior) instead of scanning every remaining rep/page just
  // to rule one out.
  var onboardingFoundAtRequest = -1;
  var ONBOARDING_ONLY_GRACE_REQUESTS = 8; // extra sequential requests to look for a demo once onboarding is found, before giving up on it
  var stopScanning = false;

  try {
    for (var repIdx = 0; repIdx < repEmails.length && !stopScanning; repIdx++) {
      var repEmail = repEmails[repIdx];
      var repPageCursor = null;

      while (true) {
        if (overallBudgetExceeded()) {
          traceLog("Fathom Fallback", "Hit the overall " + (OVERALL_SEARCH_BUDGET_MS / 1000) + "s search budget (combined across all tiers) — stopping here with whatever's been found so far", "warn");
          stopScanning = true;
          break;
        }
        if (new Date().getTime() - fbStart > fbMaxRuntimeMs) {
          traceLog("Fathom Fallback", "Stopped — hit the " + (fbMaxRuntimeMs / 1000) + "s time cap before finding both a demo and an onboarding call", "warn");
          stopScanning = true;
          break;
        }
        if (requestsMade >= fbMaxRequests) {
          traceLog("Fathom Fallback", "Stopped — hit the " + fbMaxRequests + "-request cap for this scan", "warn");
          stopScanning = true;
          break;
        }

        var seqUrl = "https://api.fathom.ai/external/v1/meetings?limit=100&include_summary=true&include_crm_matches=true"
          + "&created_after=" + encodeURIComponent(createdAfterIso)
          + (createdBeforeIso ? "&created_before=" + encodeURIComponent(createdBeforeIso) : "")
          + "&recorded_by[]=" + encodeURIComponent(repEmail)
          + (repPageCursor ? "&cursor=" + encodeURIComponent(repPageCursor) : "");

        // Pacing (elapsed-time based, uniform across rep boundaries) and
        // Retry-After handling both live in pacedFathomFetch now, shared
        // with the two tiers above — see its definition near the top of
        // this function.
        var seqData = pacedFathomFetch(seqUrl);
        requestsMade++;
        if (!seqData) {
          break; // this rep's fetch errored (or exhausted retries) — move to the next rep
        }

        fbPagesScanned++;
        var seqItems = seqData.items || [];
        fbMeetingsScanned += seqItems.length;
        for (var si2 = 0; si2 < seqItems.length; si2++) {
          if (involvesContact(seqItems[si2])) fbMatches.push(seqItems[si2]);
        }

        var hasOnboarding = fbMatches.some(function(m) { return isOnboardingMeeting(m, firstName); });
        var hasDemo = fbMatches.some(isDemoMeeting);

        if (hasOnboarding && onboardingFoundAtRequest === -1) {
          onboardingFoundAtRequest = requestsMade;
        }

        // SPEED: stop once we've found BOTH this contact's onboarding call
        // and their demo (the demo names the salesperson, and precedes
        // onboarding — so it may take a few extra requests to reach, since
        // it's older).
        if (hasOnboarding && hasDemo) {
          traceLog("Fathom Fallback", "Found both an onboarding call and a demo after " + requestsMade + " request(s) — stopping scan early", "ok");
          stopScanning = true;
          break;
        }

        // Some contacts genuinely have only ONE of the two calls
        // recorded — or reachable at all (e.g. a demo from over a year
        // ago, well outside this scan's window) — waiting on BOTH before
        // stopping just burns requests chasing something that will never
        // appear. Once an onboarding call has been found, give the demo a
        // handful of extra requests to show up; if it still hasn't, stop
        // anyway — the onboarding call alone is enough, and salesperson
        // just falls back to the deal-based guess.
        if (hasOnboarding && !hasDemo && onboardingFoundAtRequest !== -1 && (requestsMade - onboardingFoundAtRequest) >= ONBOARDING_ONLY_GRACE_REQUESTS) {
          traceLog("Fathom Fallback", "Found the onboarding call (request " + onboardingFoundAtRequest + ") but no demo turned up after " + requestsMade + " request(s) total — proceeding without a demo instead of scanning to the request/time cap. Salesperson will fall back to the deal-based guess.", "ok");
          stopScanning = true;
          break;
        }

        // No dedicated training call exists for this contact at all — a
        // demo-titled call is standing in as the real onboarding session
        // (per HubSpot's own past-demo stand-in decision). Once that
        // stand-in shows up, stop immediately instead of continuing to
        // scan for an "onboarding"-classified call that can never appear —
        // isOnboardingMeeting() will never match this contact's actual
        // onboarding call by title, so waiting for one just burns requests
        // and rate limit budget on a search that can't succeed.
        if (pastDemoStandIn && !hasOnboarding && hasDemo) {
          traceLog("Fathom Fallback", "Found the demo-titled stand-in call after " + requestsMade + " request(s) — no dedicated training call exists for this contact, so stopping here instead of continuing to scan for an 'onboarding' call that can never classify as one", "ok");
          stopScanning = true;
          break;
        }

        repPageCursor = seqData.next_cursor || null;
        if (!repPageCursor || seqItems.length === 0) break; // no more pages for this rep — move to next rep
      }
    }
  } catch (e) {
    traceLog("Fathom Fallback", "Error during scan: " + e.message + " — using whatever was gathered so far", "fail");
  }

  traceLog("Fathom Fallback", "Scanned " + fbPagesScanned + " page(s) / " + fbMeetingsScanned + " meeting(s) across " + repEmails.length + " reps (sequential, " + requestsMade + " request(s) total) — " + fbMatches.length + " matched this contact by attendee", fbMatches.length ? "info" : "warn");

  var finalResult = fbMatches.length ? summarize(fbMatches, "Fathom fallback scan") : empty;

  // ── Last-resort demo search: anchored to the DEAL's creation date ────
  // Only runs when EVERYTHING above already ran and salesperson is STILL
  // unresolved — this adds zero cost to the common case, since most
  // contacts' demo and onboarding calls happen close enough together to
  // already be found by the tiers above. Some real demos happen months
  // before the actual onboarding call (a prospect who took a while to
  // sign after their demo) — genuinely outside any onboarding- or
  // signup-anchored window we'd reasonably use for every contact. The
  // deal's OWN creation date is a real, already-known anchor for roughly
  // when that demo happened (a deal is usually created around when the
  // demo closes it), so this tries ONE narrow, targeted search around
  // THAT specific date as an extra last step — not a blanket widening of
  // the default window for everyone. Confirmed real case: Contact M's actual demo (Brock Baker) was in March, ~3 months before her
  // July onboarding — completely unreachable by any window anchored to
  // onboarding or signup, but her deal's creation date sits right next
  // to when that March demo actually happened.
  if (!finalResult.salesperson && dealCreatedDateIso && emailLc) {
    var dealDate = new Date(dealCreatedDateIso + "T00:00:00Z");
    if (!isNaN(dealDate.getTime())) {
      var dcStart = new Date(dealDate.getTime());
      dcStart.setDate(dcStart.getDate() - 14);
      var dcEnd = new Date(dealDate.getTime());
      dcEnd.setDate(dcEnd.getDate() + 14);

      traceLog("Fathom Last-Resort", "No salesperson found anywhere else — trying ONE narrow search around the deal's creation date (" + dealCreatedDateIso + "): " + isoDate(dcStart) + " to " + isoDate(dcEnd), "info");

      try {
        var lrItems = [];
        var lrCursor = null;
        for (var lp = 0; lp < 5; lp++) {
          var lrUrl = "https://api.fathom.ai/external/v1/meetings?limit=100&include_summary=true&include_crm_matches=true"
            + "&calendar_invitees[]=" + encodeURIComponent(emailLc)
            + "&created_after=" + encodeURIComponent(dcStart.toISOString())
            + "&created_before=" + encodeURIComponent(dcEnd.toISOString())
            + (lrCursor ? "&cursor=" + encodeURIComponent(lrCursor) : "");
          var lrData = pacedFathomFetch(lrUrl);
          if (!lrData) break;
          var lrPageItems = lrData.items || [];
          lrItems = lrItems.concat(lrPageItems);
          lrCursor = lrData.next_cursor || null;
          if (!lrCursor || lrPageItems.length === 0) break;
        }
        var lrVerified = lrItems.filter(involvesContact);
        var lrDemos = lrVerified.filter(isDemoMeeting);
        if (lrDemos.length) {
          lrDemos.sort(function(a, b) { return resolveTimestamp(a) - resolveTimestamp(b); });
          var lrPick = firstValidHostMeeting(lrDemos, "Salesperson (deal-date last resort)");
          if (lrPick) {
            traceLog("Fathom Last-Resort", 'Found a demo near the deal\'s creation date: "' + (lrPick.meeting.meeting_title || lrPick.meeting.title || "(untitled)") + '" — host "' + lrPick.name + '"', "ok");
            finalResult.salesperson = lrPick.name;
          } else {
            traceLog("Fathom Last-Resort", "Found demo-titled call(s) near the deal's creation date, but none had a valid-rep host — salesperson stays as the deal-based guess", "warn");
          }
        } else {
          traceLog("Fathom Last-Resort", "Nothing found near the deal's creation date either — salesperson stays as the deal-based guess", "warn");
        }
      } catch (e) {
        traceLog("Fathom Last-Resort", "Error during last-resort search: " + e.message, "fail");
      }
    }
  }

  return finalResult;
}

// ── DEBUG: see exactly what this Fathom API key can see ───────────
// Run from the Apps Script editor (select debugFathomLookup, click Run),
// then check View > Logs. This shows the raw meeting titles + recorded_by
// + dates visible to your key, and whether a name/email match is found —
// so we can tell if the issue is matching logic vs. the key simply not
// having visibility into the meeting (Fathom API keys only see meetings
// recorded by that user or shared to their team).
function debugFathomLookup() {
  // EDIT THESE: use the real name/email of a contact you know was onboarded
  var personName = "PASTE A CONTACT FULL NAME HERE";
  var personEmail = "paste-contact-email@example.com";
  var daysBack = 60; // widen this (e.g. 180) if a contact was onboarded longer ago

  var meetings;
  try {
    meetings = fetchFathomPages(daysBack);
  } catch (e) {
    Logger.log("Fathom fetch failed: " + e.message);
    return;
  }

  Logger.log("This API key can see " + meetings.length + " meeting(s) from the last " + daysBack + " days.");
  Logger.log("----------------------------------------------------");

  var nameLc = personName.toLowerCase();
  var emailLc = personEmail.toLowerCase();
  var foundMatch = false;

  for (var i = 0; i < meetings.length; i++) {
    var m = meetings[i];
    var title = m.meeting_title || m.title || "(untitled)";
    var recordedBy = (m.recorded_by && m.recorded_by.name) ? m.recorded_by.name + " <" + m.recorded_by.email + ">" : "(unknown)";
    var date = m.scheduled_start_time || m.created_at || "(no date)";
    var invitees = (m.calendar_invitees || []).map(function(c){ return c.name + " <" + c.email + ">"; }).join(", ");

    var hay = JSON.stringify(m).toLowerCase();
    var isMatch = (nameLc && hay.indexOf(nameLc) >= 0) || (emailLc && hay.indexOf(emailLc) >= 0);
    if (isMatch) foundMatch = true;

    Logger.log((isMatch ? ">>> MATCH <<<  " : "") + title + " | " + date
      + "  [onboarding call? " + (isOnboardingMeeting(m) ? "YES" : "no") + "]");
    Logger.log("   recorded_by: " + recordedBy);
    Logger.log("   invitees: " + (invitees || "(none)"));
  }

  Logger.log("----------------------------------------------------");
  if (!foundMatch) {
    Logger.log("No meeting matched \"" + personName + "\" / \"" + personEmail + "\".");
    Logger.log("If you expected a match and don't see the meeting listed above AT ALL,");
    Logger.log("the API key likely can't see it — Fathom API keys only see meetings");
    Logger.log("recorded by that key's owner, or meetings shared to their team.");
  }
}

// ── DEBUG (ALERT VERSION) — bypasses the log viewer entirely ──────
// Run this one instead if the Executions log won't show you anything.
// It pops up an on-screen alert box directly in the Google Sheet with
// the same information, so there's nothing extra to click into.
function debugFathomLookupAlert() {
  var personName = "PASTE A CONTACT FULL NAME HERE";
  var personEmail = "paste-contact-email@example.com";
  var daysBack = 60; // widen this (e.g. 180) if a contact was onboarded longer ago

  var ui = SpreadsheetApp.getUi();
  var meetings;
  try {
    meetings = fetchFathomPages(daysBack);
  } catch (e) {
    ui.alert("Fathom fetch FAILED:\n\n" + e.message);
    return;
  }

  var nameLc = personName.toLowerCase();
  var emailLc = personEmail.toLowerCase();
  var matchLines = [];
  var sampleLines = [];

  for (var i = 0; i < meetings.length; i++) {
    var m = meetings[i];
    var title = m.meeting_title || m.title || "(untitled)";
    var date = (m.scheduled_start_time || m.created_at || "(no date)").slice(0, 10);
    var recordedByName = (m.recorded_by && m.recorded_by.name) ? m.recorded_by.name : "(unknown)";
    var hay = JSON.stringify(m).toLowerCase();
    var isMatch = (nameLc && hay.indexOf(nameLc) >= 0) || (emailLc && hay.indexOf(emailLc) >= 0);
    var onb = isOnboardingMeeting(m) ? " [ONBD]" : "";

    if (isMatch) matchLines.push(title + onb + " | " + date + " | recorded by: " + recordedByName);
    if (i < 10) sampleLines.push(title + onb + " | " + date + " | recorded by: " + recordedByName);
  }

  var msg = "Total meetings visible to this key: " + meetings.length + "\n\n";

  if (matchLines.length) {
    msg += "MATCHES FOUND:\n" + matchLines.join("\n") + "\n\n";
  } else {
    msg += "NO MATCH for \"" + personName + "\" / \"" + personEmail + "\".\n\n";
  }

  msg += "First 10 meetings this key can see (for reference):\n" + (sampleLines.join("\n") || "(none — this key sees ZERO meetings)");

  ui.alert("Fathom Debug Results", msg, ui.ButtonSet.OK);
}

// ── DEBUG (LIVE SHEET VERSION) — writes progress directly into cells ──
// This is the most reliable debug option: instead of waiting for the
// whole thing to finish before showing anything (like the Alert version),
// this writes its progress into a "Fathom Debug" tab AS IT GOES. Just
// run it, then switch to that tab and watch — no waiting for a popup,
// no digging through execution logs.
function debugFathomToSheet() {
  var personName = "PASTE A CONTACT FULL NAME HERE";
  var personEmail = "paste-contact-email@example.com";
  var daysBack = 60; // widen this (e.g. 180) if a contact was onboarded longer ago

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Fathom Debug");
  if (!sheet) sheet = ss.insertSheet("Fathom Debug");
  sheet.clear();
  sheet.getRange(1, 1).setValue("Status");
  sheet.getRange(1, 2).setValue("Started at " + new Date().toLocaleTimeString() + " — checking last " + daysBack + " days");
  SpreadsheetApp.flush(); // force the write to appear immediately

  var key;
  try {
    key = getProp("FATHOM_API_KEY");
  } catch (e) {
    sheet.getRange(1, 2).setValue("ERROR: " + e.message);
    return;
  }

  var createdAfter = new Date();
  createdAfter.setDate(createdAfter.getDate() - daysBack);
  var createdAfterIso = createdAfter.toISOString();

  var all = [];
  var cursor = null;
  var startTime = new Date().getTime();
  var maxRuntimeMs = 4.5 * 60 * 1000;
  var maxPages = 500;
  var delayMs = 1100;
  var row = 3;

  sheet.getRange(row, 1).setValue("Page");
  sheet.getRange(row, 2).setValue("Meetings so far");
  sheet.getRange(row, 3).setValue("Status");
  row++;

  for (var p = 0; p < maxPages; p++) {
    if (new Date().getTime() - startTime > maxRuntimeMs) {
      sheet.getRange(1, 2).setValue("Stopped early (time cutoff) after " + p + " pages, " + all.length + " meetings.");
      SpreadsheetApp.flush();
      break;
    }

    var url = "https://api.fathom.ai/external/v1/meetings?limit=100"
      + "&created_after=" + encodeURIComponent(createdAfterIso)
      + (cursor ? "&cursor=" + encodeURIComponent(cursor) : "");
    var data = null;
    var retries = 0;
    var maxRetries = 4;

    while (retries <= maxRetries) {
      var resp = UrlFetchApp.fetch(url, {
        method: "get",
        headers: { "X-Api-Key": key },
        muteHttpExceptions: true
      });
      var code = resp.getResponseCode();

      if (code === 429) {
        var waitMs = 5000 * (retries + 1);
        sheet.getRange(row, 1).setValue(p + 1);
        sheet.getRange(row, 2).setValue(all.length);
        sheet.getRange(row, 3).setValue("429 rate limited — waiting " + (waitMs/1000) + "s (retry " + (retries+1) + "/" + maxRetries + ")");
        SpreadsheetApp.flush();
        Utilities.sleep(waitMs);
        retries++;
        continue;
      }
      if (code < 200 || code >= 300) {
        sheet.getRange(row, 1).setValue(p + 1);
        sheet.getRange(row, 3).setValue("ERROR " + code + ": " + resp.getContentText().slice(0, 200));
        SpreadsheetApp.flush();
        return;
      }
      data = JSON.parse(resp.getContentText());
      break;
    }

    if (!data) {
      sheet.getRange(row, 3).setValue("Gave up after " + maxRetries + " retries (still rate limited).");
      SpreadsheetApp.flush();
      return;
    }

    var items = data.items || [];
    all = all.concat(items);
    cursor = data.next_cursor || null;

    sheet.getRange(row, 1).setValue(p + 1);
    sheet.getRange(row, 2).setValue(all.length);
    sheet.getRange(row, 3).setValue(cursor ? "fetched, more pages remain" : "fetched, LAST page");
    SpreadsheetApp.flush(); // <-- this makes it show up live, not just at the end
    row++;

    if (!cursor || items.length === 0) break;
    Utilities.sleep(delayMs);
  }

  // ── Now write the final results ──
  row += 1;
  sheet.getRange(row, 1).setValue("DONE — total meetings visible to this key:");
  sheet.getRange(row, 2).setValue(all.length);
  row += 2;

  var nameLc = personName.toLowerCase();
  var emailLc = personEmail.toLowerCase();
  var matchCount = 0;

  sheet.getRange(row, 1).setValue("Title");
  sheet.getRange(row, 2).setValue("Date");
  sheet.getRange(row, 3).setValue("Recorded By");
  sheet.getRange(row, 4).setValue("Onboarding call?");
  sheet.getRange(row, 5).setValue("Match?");
  row++;

  for (var i = 0; i < all.length; i++) {
    var m = all[i];
    var title = m.meeting_title || m.title || "(untitled)";
    var date = (m.scheduled_start_time || m.created_at || "").slice(0, 10);
    var recordedByName = (m.recorded_by && m.recorded_by.name) ? m.recorded_by.name : "(unknown)";
    var hay = JSON.stringify(m).toLowerCase();
    var isMatch = (nameLc && hay.indexOf(nameLc) >= 0) || (emailLc && hay.indexOf(emailLc) >= 0);
    if (isMatch) matchCount++;

    sheet.getRange(row, 1).setValue(title);
    sheet.getRange(row, 2).setValue(date);
    sheet.getRange(row, 3).setValue(recordedByName);
    sheet.getRange(row, 4).setValue(isOnboardingMeeting(m) ? "YES" : "no");
    sheet.getRange(row, 5).setValue(isMatch ? "★ MATCH" : "");
    row++;
  }

  sheet.getRange(1, 2).setValue("FINISHED — " + all.length + " meetings, " + matchCount + " match(es) for \"" + personName + "\".");
  SpreadsheetApp.flush();
}

// ── Shared spreadsheet handle (avoids repeated SpreadsheetApp.openById) ──
// openById() is one of the more expensive Apps Script calls — it's a real
// network round-trip to attach the whole spreadsheet object, not a cheap
// local lookup. findExistingContactRow and findAppDataForName each used
// to call it independently, meaning every single pull paid that cost
// TWICE in a row for the exact same spreadsheet (confirmed via the trace
// timing: ~3.5s + ~1.9s back to back on a real pull). Caching the handle
// for the lifetime of one execution means the second call reuses the
// already-open object instead of re-fetching it.
var _cachedTrackingSpreadsheet = null;
function getTrackingSpreadsheet() {
  if (!_cachedTrackingSpreadsheet) {
    _cachedTrackingSpreadsheet = SpreadsheetApp.openById(SHEET_ID);
  }
  return _cachedTrackingSpreadsheet;
}

// ── Called from the frontend: pull all data for preview ───────────
// Scans the tracking sheet for an existing row matching this email
// (column B = User Email). Returns null if not found, otherwise
// { rowNumber, numCalls, firstOnboardDate } so the frontend can show
// "this is an additional call" instead of creating a duplicate row.
function findExistingContactRow(email) {
  if (!email) return null;
  var ss = getTrackingSpreadsheet();
  var sheet = ss.getSheetByName(TAB_NAME);
  if (!sheet) return null;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  var emailLc = email.trim().toLowerCase();
  var data = sheet.getRange(2, 1, lastRow - 1, 9).getValues(); // cols A-I

  for (var i = 0; i < data.length; i++) {
    var rowEmail = String(data[i][1] || "").trim().toLowerCase(); // col B
    if (rowEmail === emailLc) {
      return {
        rowNumber: i + 2, // +2 because data starts at sheet row 2, and i is 0-indexed
        firstOnboardDate: data[i][4],   // col E
        numCalls: data[i][7],           // col H
        lastCallDate: data[i][8]        // col I
      };
    }
  }
  return null;
}

// ── AppData tab lookup — sign-up date & product ────────────────────
// The "All Data - Lookup" tab (same spreadsheet) has the product each
// user is on. Columns: A=Email, B=First Name, C=Last Name, E=Products
// (confirmed directly). Matching this means the person never has to
// hand-pick the workflow — it comes straight from real data.
//
// Matches by FIRST + LAST NAME (columns B/C), not email — confirmed real
// case: Contact I's email in this tab was still his OLD employer's
// address (jthomas@leelalb.com) after he moved to Marcus & Millichap and
// HubSpot picked up his new email — an exact email match found nothing
// even though his product data was sitting right there under his
// (unchanged) name. Trade-off worth naming directly: name matching can't
// tell apart two genuinely different people who happen to share an
// identical name, the way an email address (unique per person) always
// could — a real, if likely rare, risk this change accepts in exchange
// for surviving email/employer changes, which is the more common case in
// practice for this specific tab.
//
// Sign-up date is NOT sourced from here (or anywhere automatic) anymore —
// per direct confirmation, it always matches First Onboard Date, so the
// frontend defaults it from that once known instead.
function findAppDataForName(firstName, lastName) {
  if (!firstName || !lastName) return null;
  var ss = getTrackingSpreadsheet();
  var sheet = ss.getSheetByName("All Data - Lookup");
  if (!sheet) return null;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  var firstLc = firstName.trim().toLowerCase();
  var lastLc = lastName.trim().toLowerCase();
  var data = sheet.getRange(2, 1, lastRow - 1, 5).getValues(); // cols A-E

  for (var i = 0; i < data.length; i++) {
    var rowFirst = String(data[i][1] || "").trim().toLowerCase();
    var rowLast = String(data[i][2] || "").trim().toLowerCase();
    if (rowFirst === firstLc && rowLast === lastLc) {
      var product = String(data[i][4] || "").trim();
      // Back Office and Enterprise Suite have no onboarding workflow at
      // all — treated as a non-match here, same as if nothing were found,
      // so the workflow dropdown correctly stays blank for manual
      // selection rather than auto-filling something that doesn't exist.
      var productLc = product.toLowerCase();
      if (productLc === "back office" || productLc === "enterprise suite") {
        return { product: "" };
      }
      return { product: product };
    }
  }
  return null;
}

// ── DEBUG: inspect exactly how getFathomMeetingRange picks first/last ──
// Run this to see, for a specific person, EVERY raw candidate meeting
// (before and after verification), which timestamp field was used to sort
// each one, and which ended up chosen as "first" vs "last". This is the
// fastest way to see WHY the wrong meeting got picked as the earliest —
// e.g. if one meeting has no scheduled_start_time and falls back to
// created_at, which might not sort the way you'd expect.
function debugFathomRangeToSheet() {
  var personName = "PASTE FULL NAME HERE";
  var personEmail = "paste-email@example.com";

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Fathom Range Debug");
  if (!sheet) sheet = ss.insertSheet("Fathom Range Debug");
  sheet.clear();
  sheet.getRange(1, 1).setValue("Status");
  sheet.getRange(1, 2).setValue("Running for " + personName + " / " + personEmail);
  SpreadsheetApp.flush();

  var key;
  try {
    key = getProp("FATHOM_API_KEY");
  } catch (e) {
    sheet.getRange(1, 2).setValue("ERROR: " + e.message);
    return;
  }

  var emailLc = personEmail.trim().toLowerCase();
  var nameLc = personName.toLowerCase();
  var row = 3;

  var writeHeader = function() {
    sheet.getRange(row, 1).setValue("Source");
    sheet.getRange(row, 2).setValue("Title");
    sheet.getRange(row, 3).setValue("scheduled_start_time");
    sheet.getRange(row, 4).setValue("recording_start_time");
    sheet.getRange(row, 5).setValue("created_at");
    sheet.getRange(row, 6).setValue("Field used for sort");
    sheet.getRange(row, 7).setValue("Resolved timestamp (ms)");
    sheet.getRange(row, 8).setValue("Recorded By");
    sheet.getRange(row, 9).setValue("Onboarding call?");
    sheet.getRange(row, 10).setValue("Contains email?");
    row++;
  };

  var pickField = function(m) {
    if (m.scheduled_start_time) return "scheduled_start_time";
    if (m.recording_start_time) return "recording_start_time";
    if (m.created_at) return "created_at";
    return "(none found)";
  };
  var resolveTimestamp = function(m) {
    var s = m.scheduled_start_time || m.recording_start_time || m.created_at || "";
    var t = s ? new Date(s).getTime() : NaN;
    return isNaN(t) ? 0 : t;
  };

  var writeRow = function(sourceLabel, m) {
    var containsEmail = JSON.stringify(m).toLowerCase().indexOf(emailLc) >= 0;
    sheet.getRange(row, 1).setValue(sourceLabel);
    sheet.getRange(row, 2).setValue(m.meeting_title || m.title || "(untitled)");
    sheet.getRange(row, 3).setValue(m.scheduled_start_time || "");
    sheet.getRange(row, 4).setValue(m.recording_start_time || "");
    sheet.getRange(row, 5).setValue(m.created_at || "");
    sheet.getRange(row, 6).setValue(pickField(m));
    sheet.getRange(row, 7).setValue(resolveTimestamp(m));
    sheet.getRange(row, 8).setValue(m.recorded_by && m.recorded_by.name ? m.recorded_by.name : "(unknown)");
    sheet.getRange(row, 9).setValue(isOnboardingMeeting(m) ? "YES" : "no");
    sheet.getRange(row, 10).setValue(containsEmail ? "YES" : "no");
    row++;
  };

  // ── Fetch via the email filter (primary path) ──
  sheet.getRange(1, 2).setValue("Fetching via calendar_invitees[] filter...");
  SpreadsheetApp.flush();

  var allItems = [];
  var cursor = null;
  for (var p = 0; p < 20; p++) {
    var url = "https://api.fathom.ai/external/v1/meetings?limit=100"
      + "&calendar_invitees[]=" + encodeURIComponent(emailLc)
      + (cursor ? "&cursor=" + encodeURIComponent(cursor) : "");
    var resp = UrlFetchApp.fetch(url, { method: "get", headers: { "X-Api-Key": key }, muteHttpExceptions: true });
    if (resp.getResponseCode() < 200 || resp.getResponseCode() >= 300) break;
    var data = JSON.parse(resp.getContentText());
    var items = data.items || [];
    allItems = allItems.concat(items);
    cursor = data.next_cursor || null;
    if (!cursor || items.length === 0) break;
    Utilities.sleep(600);
  }

  row += 1;
  sheet.getRange(row, 1).setValue("RAW results from calendar_invitees[] filter: " + allItems.length + " meeting(s)");
  row += 1;
  writeHeader();
  for (var i = 0; i < allItems.length; i++) writeRow("filter-raw", allItems[i]);

  var verified = allItems.filter(function(m) {
    return JSON.stringify(m).toLowerCase().indexOf(emailLc) >= 0;
  });

  row += 1;
  sheet.getRange(row, 1).setValue("VERIFIED (actually contain the email): " + verified.length + " meeting(s)");
  row += 1;

  // Onboarding-only view — this is what actually feeds the count + dates.
  var onboardingsOnly = verified.filter(isOnboardingMeeting);
  row += 1;
  sheet.getRange(row, 1).setValue("ONBOARDING CALLS ONLY (demos excluded): " + onboardingsOnly.length + " meeting(s)");
  row += 1;

  if (onboardingsOnly.length) {
    writeHeader();
    var sorted = onboardingsOnly.slice().sort(function(a, b) { return resolveTimestamp(a) - resolveTimestamp(b); });
    for (var j = 0; j < sorted.length; j++) writeRow("onboarding-sorted[" + j + "]", sorted[j]);

    row += 1;
    sheet.getRange(row, 1).setValue("=> FIRST (index 0):");
    sheet.getRange(row, 2).setValue(sorted[0].meeting_title || sorted[0].title || "(untitled)");
    sheet.getRange(row, 3).setValue(isoDate(new Date(resolveTimestamp(sorted[0]))));
    row++;
    sheet.getRange(row, 1).setValue("=> LAST (index " + (sorted.length - 1) + "):");
    sheet.getRange(row, 2).setValue(sorted[sorted.length - 1].meeting_title || sorted[sorted.length - 1].title || "(untitled)");
    sheet.getRange(row, 3).setValue(isoDate(new Date(resolveTimestamp(sorted[sorted.length - 1]))));
  } else {
    sheet.getRange(row, 1).setValue("No onboarding calls found — would fall back to the name/email scan, or return empty.");
  }

  sheet.getRange(1, 2).setValue("FINISHED. See rows below.");
  SpreadsheetApp.flush();
}

// ── DEBUG: test the existing-row detection logic in isolation ─────
// Run this directly to check whether a specific email correctly matches
// (or doesn't match) a row already in the sheet — without needing to go
// through the full HubSpot + Fathom pull. Pops up an alert with either
// the exact row it found, or a dump of every email currently in column B
// so you can visually compare and spot a mismatch (extra space, typo,
// different casing, etc).
function debugCheckExistingRow() {
  var testEmail = "PASTE THE EMAIL TO TEST HERE";

  var ui = SpreadsheetApp.getUi();
  var ss = getTrackingSpreadsheet();
  var sheet = ss.getSheetByName(TAB_NAME);
  if (!sheet) {
    ui.alert('Tab "' + TAB_NAME + '" not found in this spreadsheet.');
    return;
  }

  var result = findExistingContactRow(testEmail);

  var lastRow = sheet.getLastRow();
  var allEmails = [];
  if (lastRow >= 2) {
    var data = sheet.getRange(2, 2, lastRow - 1, 1).getValues(); // col B only
    for (var i = 0; i < data.length; i++) {
      var e = String(data[i][0] || "").trim();
      if (e) allEmails.push("Row " + (i + 2) + ": \"" + e + "\"");
    }
  }

  var msg;
  if (result) {
    msg = "✅ MATCH FOUND\n\n"
      + "Testing email: \"" + testEmail + "\"\n"
      + "Matched row: " + result.rowNumber + "\n"
      + "First Onboard Date on file: " + result.firstOnboardDate + "\n"
      + "Current # of calls: " + result.numCalls + "\n"
      + "Current Last Call Date: " + result.lastCallDate;
  } else {
    msg = "❌ NO MATCH\n\n"
      + "Testing email: \"" + testEmail + "\"\n\n"
      + "Every email currently in column B (" + allEmails.length + " row(s)):\n"
      + (allEmails.join("\n") || "(sheet appears empty below the header)")
      + "\n\nCompare closely for: extra spaces, different casing (matching is case-insensitive so that's NOT it), "
      + "or the email actually being slightly different (e.g. a personal vs work email).";
  }

  ui.alert("Existing Row Check", msg, ui.ButtonSet.OK);
}

// ── Called from the frontend: live name search for the search dropdown ──
// Searches HubSpot contacts by first name, last name, or email containing
// any of the typed words. Returns a short list for the person to click,
// so there's no need to go find and paste a HubSpot URL at all.
function searchContactsByName(query) {
  if (!query) return [];
  var trimmed = query.trim();
  if (!trimmed) return [];

  try {
    var filterGroups;

    if (trimmed.indexOf("@") >= 0) {
      // Looks like an email — one simple, fast filter instead of the
      // word-splitting approach below. Wildcard-wrapped so a partial
      // email (e.g. "josh.kirby@stanb") matches too, not just a complete
      // token.
      filterGroups = [{ filters: [{ propertyName: "email", operator: "CONTAINS_TOKEN", value: "*" + trimmed + "*" }] }];
    } else {
      // Name-style query — check firstname/lastname per word (capped at
      // 3 words; nobody's searching a 5-word name). Deliberately skips
      // the email field here since matching name tokens against email
      // addresses rarely helps and only adds filter groups for HubSpot
      // to evaluate, which is where a chunk of the latency comes from.
      var words = trimmed.split(/\s+/).filter(function(w) { return w.length > 0; }).slice(0, 3);
      if (!words.length) return [];

      if (words.length === 1) {
        // Wildcard-wrapped so a partial name (e.g. "dodso") correctly
        // prefix-matches a complete one ("Dodson") — CONTAINS_TOKEN
        // without wildcards only matches a COMPLETE token, so "dodso"
        // would never match "dodson" at all without this, no matter how
        // the search got triggered (typing further, or deleting back
        // down to a shorter partial word).
        filterGroups = [
          { filters: [{ propertyName: "firstname", operator: "CONTAINS_TOKEN", value: "*" + words[0] + "*" }] },
          { filters: [{ propertyName: "lastname", operator: "CONTAINS_TOKEN", value: "*" + words[0] + "*" }] }
        ];
      } else {
        // Multi-word: require EVERY word to appear SOMEWHERE (in some
        // field), not just ANY ONE word in ANY ONE field. HubSpot's search
        // ORs across separate filterGroups but ANDs filters WITHIN one
        // group — so getting "every word matches somewhere, in any field
        // assignment" means enumerating every way the words could map onto
        // firstname/lastname (the cross product), with each full mapping
        // as its own AND'd group, then OR-ing all of those mappings
        // together. An earlier version pushed one independent filterGroup
        // per (word, field) pair — which HubSpot ORs, meaning a query for
        // "Liz Richardson" matched ANY contact with "Richardson" as a last
        // name ALONE. With 112 real Richardsons in the database and a
        // limit of 10, her actual combined match never had a chance to
        // surface — confirmed as the exact cause of two real reported
        // failures (Liz Richardson, Contact L), not a coincidence.
        var fieldOptions = ["firstname", "lastname"];
        var combos = [[]];
        words.forEach(function(w) {
          var next = [];
          combos.forEach(function(combo) {
            fieldOptions.forEach(function(f) {
              next.push(combo.concat([{ field: f, word: w }]));
            });
          });
          combos = next;
        });
        filterGroups = combos.map(function(combo) {
          return { filters: combo.map(function(c) { return { propertyName: c.field, operator: "CONTAINS_TOKEN", value: "*" + c.word + "*" }; }) };
        });
      }
    }

    var data = hsFetch("/crm/v3/objects/contacts/search", "post", {
      filterGroups: filterGroups,
      properties: ["firstname", "lastname", "email", "company"],
      limit: 10
    });

    var results = data.results || [];
    return results.map(function(c) {
      var p = c.properties || {};
      return {
        id: c.id,
        name: [p.firstname, p.lastname].filter(function(x) { return x; }).join(" ") || "(no name)",
        email: p.email || "",
        company: p.company || ""
      };
    });
  } catch (e) {
    return []; // fail quietly — this just means the dropdown shows no results
  }
}

// ── AI Accuracy Checker ──────────────────────────────────────────────
// Runs a real AI (Claude) second-opinion pass on Lupe's own classification
// and rep-assignment logic — the deterministic trace shows WHICH rule
// fired, but can't judge whether the underlying call actually made sense.
// This gives that judgment a chance to catch things a keyword/regex match
// can't, e.g. a title that technically matched "onboard" but doesn't
// really read like a genuine onboarding session once you see the summary.
//
// Runs automatically after every pull, alongside the trace. Never blocks
// or fails the pull — if the API key is missing, the call errors, or the
// response can't be parsed, it just reports itself as unavailable and the
// rest of the row still loads normally.
//
// Requires a Script Property ANTHROPIC_API_KEY (same pattern as
// HUBSPOT_TOKEN / FATHOM_API_KEY) — get a key at console.anthropic.com.
function getAiAccuracyCheck(ctx) {
  // Master switch — when off, never even attempt a network call. This is
  // the only paid part of Lupe; everything else keeps working regardless.
  if (!AI_ACCURACY_CHECK_ENABLED) {
    return { available: false, reason: "AI accuracy check is turned off (AI_ACCURACY_CHECK_ENABLED = false in Code.gs). Lupe runs entirely free with this disabled." };
  }

  // Nothing to check — don't spend an API call on a contact where Lupe
  // itself found no title/rep evidence at all.
  if (!ctx.demoTitle && !ctx.onboardingTitle && !ctx.salesperson && !ctx.onboarder) {
    return { available: false, reason: "Nothing to verify — no classification or rep data was found for this contact." };
  }

  var key;
  try {
    key = getProp("ANTHROPIC_API_KEY");
  } catch (e) {
    return { available: false, reason: "No ANTHROPIC_API_KEY Script Property set — add one from console.anthropic.com to enable this." };
  }

  var prompt = buildAiCheckPrompt(ctx);

  try {
    var resp = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
      method: "post",
      headers: {
        "x-api-key": key,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json"
      },
      payload: JSON.stringify({
        model: "claude-sonnet-4-6",
        max_tokens: 700,
        messages: [{ role: "user", content: prompt }]
      }),
      muteHttpExceptions: true
    });

    var code = resp.getResponseCode();
    if (code < 200 || code >= 300) {
      return { available: false, reason: "Anthropic API " + code + ": " + resp.getContentText().slice(0, 200) };
    }

    var data = JSON.parse(resp.getContentText());
    var text = (data.content || []).map(function(b) { return b.text || ""; }).join("");
    var parsed = parseAiJson(text);
    if (!parsed) {
      return { available: false, reason: "Could not parse the AI's response as JSON." };
    }
    parsed.available = true;
    return parsed;
  } catch (e) {
    return { available: false, reason: "AI check failed: " + e.message };
  }
}

function buildAiCheckPrompt(ctx) {
  return [
    "You are auditing an automated onboarding-tracking tool called Lupe. It just resolved data for a customer contact using keyword-matching rules on Fathom call titles and a fixed roster of valid reps. Your job is ONLY to sanity-check its two riskiest judgment calls. You have no access to full call transcripts — only titles, hosts, dates, and (when available) a short AI-generated call summary.",
    "",
    "VALID REPS (the only people who can legitimately be a salesperson or onboarder): " + VALID_REPS.join(", "),
    "",
    "CONTACT: " + ctx.contactName + " (" + (ctx.organization || "unknown org") + ")",
    "",
    "DEMO CALL Lupe selected (used to determine the SALESPERSON):",
    '  Title: "' + (ctx.demoTitle || "(none found)") + '"',
    "  Host Lupe assigned as salesperson: " + (ctx.salesperson || "(none)"),
    (ctx.demoSummary ? "  Fathom call summary: " + ctx.demoSummary : "  (no call summary available — judge on title alone)"),
    "",
    "ONBOARDING/TRAINING CALL Lupe selected (used to determine the ONBOARDER):",
    '  Title: "' + (ctx.onboardingTitle || "(none found)") + '"',
    "  Host Lupe assigned as onboarder: " + (ctx.onboarder || "(none)"),
    (ctx.onboardingSummary ? "  Fathom call summary: " + ctx.onboardingSummary : "  (no call summary available — judge on title alone)"),
    "",
    "Answer two questions:",
    "1. CLASSIFICATION: Does the demo call genuinely sound like an initial sales demo, and does the onboarding call genuinely sound like an onboarding/training session — not a demo, and not a later touchpoint like a renewal, refund, or check-in call?",
    "2. REP ASSIGNMENT: Does anything in the evidence suggest the assigned salesperson or onboarder is wrong — e.g. the host isn't actually one of the valid reps, or the summary suggests a different person actually ran the call?",
    "",
    "Respond with ONLY this JSON object and nothing else — no markdown fences, no preamble:",
    '{"classificationConfidence":"high|medium|low","classificationNotes":"...","repConfidence":"high|medium|low","repNotes":"...","flags":["..."]}',
    'If everything looks fine, use "high" confidence and an empty flags array. Keep each notes field to one short sentence.'
  ].join("\n");
}

// Parses Claude's JSON response defensively — strips markdown code fences
// if present, and falls back to extracting the first {...} block if the
// model added any stray text around the JSON.
function parseAiJson(text) {
  if (!text) return null;
  var cleaned = text.replace(/```json/gi, "").replace(/```/g, "").trim();
  try {
    return JSON.parse(cleaned);
  } catch (e) {
    var match = cleaned.match(/\{[\s\S]*\}/);
    if (match) {
      try { return JSON.parse(match[0]); } catch (e2) { return null; }
    }
    return null;
  }
}

function pullContactData(input) {
  TRACE = []; // reset the accuracy trace for this pull
  TRACE_START_MS = new Date().getTime();

  var contactId = extractContactId(input);
  var isEmail = !contactId && String(input).indexOf("@") >= 0;

  if (!contactId && !isEmail) {
    traceLog("Input", 'Could not extract a contact ID or email from: "' + input + '"', "fail");
    return { ok: false, error: "Couldn't find a contact ID in that URL. Paste the HubSpot record URL, a contact ID, or an email.", trace: TRACE };
  }
  traceLog("Input", isEmail ? 'Interpreted as an EMAIL: "' + input + '"' : "Interpreted as a HubSpot contact ID: " + contactId, "info");

  try {
    if (isEmail) {
      var search = hsFetch("/crm/v3/objects/contacts/search", "post", {
        filterGroups: [{ filters: [{ propertyName: "email", operator: "EQ", value: String(input).trim() }] }],
        properties: ["firstname","lastname","email","company","hubspot_owner_id"],
        limit: 1
      });
      if (!search.results || !search.results.length) {
        traceLog("Input", "No HubSpot contact found matching that email", "fail");
        return { ok: false, error: "No HubSpot contact found with that email.", trace: TRACE };
      }
      contactId = search.results[0].id;
      traceLog("Input", "Email resolved to contact ID " + contactId, "ok");
    }

    var contact = getHubSpotContact(contactId);
    traceLog("HubSpot Contact", 'Fetched contact ' + contactId + ' — firstname="' + contact.firstName + '", lastname="' + contact.lastName + '", company="' + contact.company + '"', "info");

    // Smart resolution: company falls back to the associated company record
    // when the flat property is blank; name strips out a company that leaked
    // into the name fields, or reconstructs from a dotted email local-part.
    var organization = resolveCompanyName(contact, contactId);
    var fullName = resolveContactName(contact, organization);

    // Fathom's UI has no direct per-recording deep link from the API, but
    // it does support searching by CONTACT (email), which matches the
    // real attendee record regardless of what the call is titled —
    // confirmed far more reliable than the title-text search this used
    // to use: Contact J's real calls are titled "...with Collin" and
    // "...with Devyn" (owner/first-name only), never containing her own
    // last name at all, so a title search for "Brunet" correctly finds
    // nothing even though her recordings genuinely exist. This is the
    // same pattern seen repeatedly throughout this project — a customer's
    // last name is frequently absent from the call title — so search-by-
    // title was never a reliable way to open a specific person's
    // recordings, only a coincidental one. Search-by-contact-email
    // doesn't have that problem since it doesn't depend on title wording
    // at all.
    // encodeURIComponent turns "@" into "%40", but confirmed in practice
    // that Fathom's own routing doesn't decode that back before matching
    // the path — the encoded form 404s while the exact same URL with a
    // literal "@" works fine. "@" is actually a character path segments
    // are allowed to contain unencoded per URL spec (RFC 3986 pchar), so
    // encoding it here was unnecessary in the first place; still safely
    // encoding everything else in the email, just restoring the literal
    // "@" afterward to match what Fathom's router actually expects.
    var fathomSearchUrl = contact.email
      ? "https://fathom.video/calls/search/contact/" + encodeURIComponent(contact.email).replace(/%40/g, "@")
      : "";

    var deal = getContactDeal(contactId);
    // Referral always defaults to N — per direct confirmation from the
    // onboarders, deal Source is NOT a reliable signal for whether this
    // specific customer was personally referred (it tracks deal/sales
    // attribution, a different concept). Confirmed by this exact rule
    // being wrong on 4 separate real contacts this session (Devyn, Krish,
    // Contact L, and implicitly Contact M) — every one had a deal source
    // matching the old keyword list, and every one was actually N on the
    // real tracking sheet. Still manually editable in the dropdown if a
    // rep knows a specific customer genuinely was referred.
    var isReferral = "N";
    traceLog("Referral", "Always defaults to N — not auto-detected from deal source (per onboarders' confirmation). Edit manually if this specific customer was genuinely referred.", "info");

    // Fast, reliable rep source: the contact's Meetings tab (demo host =
    // salesperson, training host = onboarder), validated against the roster.
    // This alone is usually enough to name the reps WITHOUT ever touching
    // Fathom — a couple of quick HubSpot calls, not a multi-page scan.
    var hsReps = getMeetingRepsFromHubSpot(contactId, fullName);

    // Deal owner (or, if named directly, the deal Source field) is a
    // TRUSTED source for salesperson — confirmed directly by the
    // onboarders that HubSpot's deal owner is the intended source of
    // truth for this field. A genuine HubSpot Meetings-tab record (an
    // actual demo call, not just a deal assignment) still takes priority
    // when one exists, since it's an even more direct signal of who
    // actually ran the sales call — but the deal is no longer excluded
    // as an unverified guess; it's a legitimate fallback now, exactly
    // the way onboarder already worked.
    var validatedDealSalesperson = validRep(deal.salesperson);
    var salesperson = hsReps.salesperson || validatedDealSalesperson || "";
    var onboarder   = hsReps.onboarder || "";
    var salespersonSourceLabel = hsReps.salesperson ? "HubSpot Meetings tab" : (validatedDealSalesperson ? "Deal owner/source" : "none yet");
    traceLog("Final: Salesperson (fast path)", 'Selected "' + (salesperson || "(pending Fathom)") + '" — source: ' + salespersonSourceLabel, salesperson ? "ok" : "info");
    traceLog("Final: Onboarder (fast path)", 'Selected "' + (onboarder || "(pending Fathom)") + '" — source: ' +
      (hsReps.onboarder ? "HubSpot Meetings tab" : "none yet"), onboarder ? "ok" : "info");

    // Check if this person is already in the tracking sheet.
    var existing = findExistingContactRow(contact.email);
    traceLog("Tracking Sheet", existing
      ? "Existing row found (row " + existing.rowNumber + ") — will be treated as a repeat contact"
      : "No existing row found — will insert a new row", "info");

    // Look up real sign-up date + product from the AppData tab, so the
    // person doesn't have to manually type either one. Matched by name,
    // not email — see findAppDataForName's comment for why.
    var appData = findAppDataForName(contact.firstName, contact.lastName);
    if (appData) {
      traceLog("AppData", 'Matched by name in "All Data - Lookup" — product: "' + (appData.product || "(none / ignored)") + '"', "ok");
    } else if (deal.dealProducts) {
      // Fallback to HubSpot's own deal Product(s) field when the sheet
      // has no row for this contact at all. The frontend's
      // mapProductToWorkflow already handles the semicolon-separated
      // format this field uses (e.g. "OS Prospects;Deal View") and
      // correctly infers Suite when both Prospects and Deal View are
      // present together, matching the confirmed rule that Suite isn't a
      // literal selectable value anywhere — it's always this bundle.
      appData = { product: deal.dealProducts };
      traceLog("AppData", 'No match in "All Data - Lookup" — falling back to the deal\'s own Product(s) field: "' + deal.dealProducts + '"', "ok");
    } else {
      traceLog("AppData", "No AppData row matched this name, and the deal has no Product(s) set either — sign-up date and workflow must be entered manually", "warn");
    }

    // Fathom (onboard/last-call dates, call count, AI-check evidence) is
    // now resolved in a SEPARATE async call the frontend fires right after
    // this one returns — see resolveFathomAndFinalize. This response has
    // everything HubSpot alone can answer, which is almost always enough to
    // start reviewing the contact; date fields below are placeholders the
    // frontend fills in once Fathom actually finishes.
    traceLog("Fathom", "Resolving asynchronously — dates, call count, and the AI check will fill in shortly.", "info");

    var priorCalls = existing ? parseInt(existing.numCalls, 10) : 0;
    if (isNaN(priorCalls)) priorCalls = 0;

    var rowData = {
      userName: fullName,
      userEmail: contact.email,
      organization: organization,
      userType: deal.suggestedUserTypeFromDealSize || "Individual",
      firstOnboardDate: existing ? (existing.firstOnboardDate || "") : "", // filled in by the async Fathom step for new contacts
      salesperson: salesperson,
      onboarder: onboarder,
      numCalls: existing ? String(priorCalls) : "", // repeat-contact bump / true count both need Fathom — set async
      lastCallDate: "",
      referral: isReferral,
      testimonial: "N"
    };

    return {
      ok: true,
      contactId: contactId,
      row: rowData,
      dealSource: deal.source || "",
      existingRow: existing ? existing.rowNumber : null, // null = new contact, number = updating this row
      appData: appData, // { product } or null if no match in All Data - Lookup
      fathomSearchUrl: fathomSearchUrl, // e.g. https://fathom.video/calls/search/contact/stephen.scuderi@example.com
      trace: TRACE, // step-by-step resolution log for the accuracy sidebar
      fathomPending: true, // tells the frontend to call resolveFathomAndFinalize next
      fathomContext: { fullName: fullName, email: contact.email, organization: organization, dealSource: deal.source || "",
                        hasHsSalesperson: !!hsReps.salesperson, hasHsOnboarder: !!hsReps.onboarder,
                        existingRow: existing ? existing.rowNumber : null, existingFirstOnboardDate: existing ? (existing.firstOnboardDate || "") : "",
                        existingPriorCalls: priorCalls, hsSalesperson: salesperson, hsOnboarder: onboarder,
                        hsSalespersonSource: salespersonSourceLabel,
                        onboardingScheduledFuture: !!hsReps.onboardingScheduledFuture,
                        onboardingScheduledDateIso: hsReps.onboardingScheduledDateIso || "",
                        onboarderFromPastDemo: !!hsReps.onboarderFromPastDemo,
                        pastDemoDateIso: hsReps.pastDemoDateIso || "",
                        candidateOnboardingDateIso: hsReps.candidateOnboardingDateIso || "",
                        candidateLastOnboardingDateIso: hsReps.candidateLastOnboardingDateIso || "",
                        dealCreatedDateIso: deal.createdDateIso || "",
                        suggestedUserTypeFromDealSize: deal.suggestedUserTypeFromDealSize || "",
                        knownPastTrainingCount: hsReps.knownPastTrainingCount || 0 }
    };
  } catch (e) {
    traceLog("Error", e.message || String(e), "fail");
    return { ok: false, error: e.message || String(e), trace: TRACE };
  }
}

// ── Called from the frontend: the async second half of a pull ─────────
// Runs the slow Fathom resolution (cached when possible) and the AI check,
// AFTER pullContactData has already returned and rendered everything else.
// Returns just the pieces that depend on Fathom: dates, call count, the
// final (Fathom-corrected) salesperson/onboarder, and the AI check context.
function resolveFathomAndFinalize(contactId, ctx) {
  // Reset the GLOBAL trace for this call. getFathomMeetingRange and
  // everything it calls (the Window/Invitee-Filter/Fallback/Classify
  // logging) all write to this same global TRACE via traceLog() — a
  // previous version of this function kept its own separate local trace
  // array instead, which meant every bit of that diagnostic detail was
  // being generated correctly but silently discarded before ever reaching
  // the frontend. Using the same global trace here (and returning it
  // directly) is what actually surfaces it.
  TRACE = [];
  TRACE_START_MS = new Date().getTime();
  var log = traceLog; // same (step, detail, status) signature used below

  // Defensive: if ctx ever arrives missing or malformed (e.g. a frontend/
  // backend version mismatch during a deploy), fail gracefully with a clear
  // error instead of throwing a raw TypeError — the frontend's failure
  // handler can then reveal the row with a warning instead of hanging
  // indefinitely waiting for a response that already crashed.
  if (!ctx || typeof ctx !== "object") {
    return { ok: false, error: "resolveFathomAndFinalize received no context object — this usually means Code.gs and Index.html are on mismatched versions. Re-paste both files fresh and redeploy." };
  }
  ctx.fullName = ctx.fullName || "";
  ctx.email = ctx.email || "";
  ctx.organization = ctx.organization || "";

  // A past "OneSource Meeting"/demo call is standing in for the
  // onboarding session because HubSpot's own Meetings tab has no
  // dedicated training call for this contact. This USED to short-circuit
  // straight to using that date, skipping Fathom entirely — that was
  // wrong: HubSpot's Meetings tab can be genuinely incomplete, not just
  // imprecise. Confirmed real case: Contact N and Contact O each had 4 real, dated onboarding recordings in Fathom —
  // logged under a THIRD person's calendar (a team coordinator helping
  // with setup), so neither contact's own HubSpot Meetings tab ever knew
  // these sessions existed. Skipping the Fathom check here would have
  // permanently missed them, reporting a months-old demo date as "first
  // onboard date" instead of the real, much later onboarding activity.
  // Fathom is now always consulted below — see the new past-demo-anchored
  // search window in getFathomMeetingRange, and the final date/count
  // computation further down, which only falls back to this stand-in
  // date when Fathom's search genuinely finds nothing at all.

  // NOTE: an earlier version of this function had a shortcut here that
  // skipped the Fathom scan entirely whenever HubSpot's Meetings tab
  // already had a validated onboarder AND salesperson, on the theory that
  // a "confirmed" HubSpot record made Fathom redundant. That was wrong: it
  // trusted HubSpot's meeting OWNER field as good enough on its own, which
  // directly contradicts this project's core, hard-won rule — HubSpot's
  // meeting owner is often just whoever the record is filed under, not who
  // actually ran the call, which is exactly why Fathom's recorded_by
  // (checked below) is the authoritative source for who the onboarder
  // really was. That shortcut also used every training-titled HubSpot
  // meeting as the call count, including ones scheduled in the future that
  // haven't actually happened (and so have no Fathom recording at all) —
  // overcounting completed onboarding calls. Fathom is always consulted
  // below; its recorded_by can and does override the HubSpot fast-path
  // guess (see "Final: Onboarder" / "Final: Salesperson" further down),
  // and its call count only reflects calls that actually have a recording.

  // A future-scheduled training call on HubSpot's Meetings tab used to
  // short-circuit straight to skipping Fathom entirely — reasoning that
  // Fathom can't possibly have a recording of a call that hasn't happened
  // yet. That's true for THAT specific scheduled call, but wrong as a
  // reason to give up on Fathom altogether: HubSpot's title-based
  // classification can simply miss a DIFFERENT, already-happened real
  // onboarding call, the same way it missed Contact G's and Contact N /
  // Contact O's. Confirmed real case: Contact K had a genuine "STE ONBD"-
  // tagged call TODAY — his profile shows 3 meetings total but only 2 get
  // classified (1 demo, 1 future training), leaving a third, unclassified
  // meeting unaccounted for, almost certainly today's real session under
  // a title HubSpot's own logic didn't recognize. Fathom is now always
  // consulted — with no other date anchor available here, this falls
  // through naturally to the existing "quick guess" (today ± 1 day)
  // window below, which is exactly the right net for a call that just
  // happened very recently.

  var cacheKey = "lupe_fathom_" + CODE_VERSION + "_" + contactId;
  var cached = getCachedFathomResolution(cacheKey);
  var fathom;
  if (cached) {
    log("Cache", "Reusing Fathom resolution from the last " + PULL_CACHE_TTL_SECONDS + "s — skipped the Fathom scan entirely.", "ok");
    fathom = cached.fathom;
  } else {
    fathom = getFathomMeetingRange(ctx.fullName, ctx.email, ctx.organization, ctx.signupDateIso, ctx.candidateOnboardingDateIso, ctx.candidateLastOnboardingDateIso, ctx.dealCreatedDateIso, ctx.knownPastTrainingCount, ctx.pastDemoDateIso, ctx.suggestedUserTypeFromDealSize, ctx.hsOnboarder, !!ctx.onboarderFromPastDemo);
    // Reuses ctx.hsSalesperson/ctx.hsOnboarder (already fetched once in the
    // fast path) instead of calling getMeetingRepsFromHubSpot a second
    // time — that redundant re-fetch was adding 2 avoidable HubSpot calls
    // to every single pull's async step for no new information.
    //
    // Only cache a GENUINE find (an onboarding call was actually located).
    // Caching an empty/failed result would mean a scan that came back
    // empty — whether from a real miss or from hitting a cap while still
    // legitimately working — gets replayed instantly for 5 minutes,
    // masking the fact that a fresh attempt might succeed. Every retry on
    // a contact Fathom hasn't found yet should actually retry, not reuse
    // a cached miss.
    if (fathom && fathom.totalCount > 0) {
      setCachedFathomResolution(cacheKey, { fathom: fathom });
    } else {
      log("Cache", "Not caching this result — no onboarding call was found, so the next pull will retry the scan fresh instead of reusing a miss.", "info");
    }
  }

  // Fathom's demo/training host, when available, supersedes the fast-path
  // guess already shown to the rep (HubSpot Meetings tab, or the deal
  // owner/source as a trusted fallback — confirmed directly by the
  // onboarders that deal owner is a legitimate source for this field).
  var salesperson = fathom.salesperson || ctx.hsSalesperson || "";
  var onboarder = fathom.firstOnboarder || ctx.hsOnboarder || "";
  log("Final: Salesperson", 'Selected "' + (salesperson || "(blank)") + '" — source: ' +
    (fathom.salesperson ? "Fathom demo host" : ctx.hsSalesperson ? (ctx.hsSalespersonSource || "HubSpot Meetings tab") : "none — left blank"),
    salesperson ? "ok" : "warn");
  log("Final: Onboarder", 'Selected "' + (onboarder || "(blank)") + '" — source: ' +
    (fathom.firstOnboarder ? "Fathom training host" : ctx.hsOnboarder ? "HubSpot Meetings tab" : "none — left blank"), onboarder ? "ok" : "warn");

  var numCalls, firstOnboardDate, lastCallDate;
  if (ctx.existingRow) {
    var priorCalls = ctx.existingPriorCalls || 0;
    numCalls = String(priorCalls + 1);
    firstOnboardDate = ctx.existingFirstOnboardDate || earlierIsoDate(fathom.firstDate, ctx.candidateOnboardingDateIso) || ctx.pastDemoDateIso || "";
    lastCallDate = laterIsoDate(fathom.lastDate, ctx.candidateLastOnboardingDateIso) || ctx.pastDemoDateIso || "";
  } else {
    // The date RANGE and COUNT use the widest/highest known across BOTH
    // sources, not just whichever one happened to search first — a
    // meeting's scheduled date and the fact that it occurred at all are
    // reliable basic facts straight from HubSpot's own calendar, not
    // something that genuinely needs Fathom's verification the way HOST
    // attribution does (that part still only ever trusts Fathom's
    // recorded_by, per the onboarder/salesperson logic above). Confirmed
    // real gap this closes: Fathom finding SOME but not ALL of the known
    // real calls (Contact I — verified only 1 of 3 "Ramp Up Call"
    // sessions before a rate-limit wall cut the search short) used to let
    // that partial result win outright, since the earlier fallback only
    // applied when Fathom found literally zero — reporting 6/9 (the one
    // call it reached) instead of the true 6/1-through-6/9 range HubSpot's
    // own calendar already knew with confidence.
    //
    // ctx.pastDemoDateIso is the ULTIMATE fallback — used only when
    // NEITHER Fathom NOR a real HubSpot known-training-date has anything
    // at all. This restores the original "use the past-demo stand-in"
    // behavior for genuinely quiet contacts, but now only as a true last
    // resort AFTER actually searching Fathom (per the new past-demo-
    // anchored window above), not a shortcut that skipped the search
    // entirely and could silently miss real recordings HubSpot's own
    // Meetings tab never knew existed (confirmed real case: Contact N / Contact O — 4 real sessions logged under a third
    // person's calendar; Contact G — a real "STE ONBD"-tagged call
    // titled generically enough to slip past HubSpot's own title logic
    // entirely).
    numCalls = String(Math.max(fathom.totalCount || 0, ctx.knownPastTrainingCount || 0));
    firstOnboardDate = earlierIsoDate(fathom.firstDate, ctx.candidateOnboardingDateIso) || ctx.pastDemoDateIso || "";
    lastCallDate = laterIsoDate(fathom.lastDate || fathom.firstDate, ctx.candidateLastOnboardingDateIso || ctx.candidateOnboardingDateIso) || ctx.pastDemoDateIso || "";
  }

  return {
    ok: true,
    salesperson: salesperson,
    onboarder: onboarder,
    numCalls: numCalls,
    firstOnboardDate: firstOnboardDate,
    lastCallDate: lastCallDate,
    // Suggested from the real onboarding call's own attendee list — more
    // than one non-CRE-OneSource attendee means Team, otherwise
    // Individual. Blank when no attendee data was available at all (e.g.
    // Fathom found nothing, or the meeting has no linked calendar event) —
    // the frontend leaves the dropdown at its existing default in that
    // case rather than forcing a guess.
    suggestedUserType: fathom.suggestedUserType || "",
    trace: TRACE,
    aiCheckContext: {
      contactName: ctx.fullName,
      organization: ctx.organization,
      demoTitle: fathom.demoTitle || "",
      demoSummary: fathom.demoSummary || "",
      salesperson: salesperson,
      onboardingTitle: fathom.onboardingTitle || "",
      onboardingSummary: fathom.onboardingSummary || "",
      onboarder: onboarder
    }
  };
}

// ── Called from the frontend: write the row (new contact OR repeat call) ──
// If targetRow is provided, this UPDATES that existing row's # of calls
// and last call date instead of inserting a brand new row — used when
// findExistingContactRow() already found this person in the sheet.
function insertRow(row, targetRow) {
  // Script-wide lock guarding the sheet write itself — two onboarders
  // submitting at nearly the same moment used to be a real race: both
  // could read "insert at row 2" and step on each other, or one person's
  // insertRowBefore(2) could shift row numbers out from under an update
  // already in flight for someone else's targetRow. Waits up to 30s for
  // the lock; if it can't get one in that window, fails loudly with a
  // clear "try again" message instead of silently corrupting a row —
  // this only ever serializes the brief moment of the actual write, not
  // the whole HubSpot/Fathom pull that happens before it.
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(30000)) {
      return { ok: false, error: "Sheet is busy (another onboarder is saving right now) — please wait a moment and try again." };
    }
  } catch (e) {
    return { ok: false, error: "Couldn't acquire sheet lock: " + (e.message || String(e)) };
  }

  try {
    var ss = getTrackingSpreadsheet();
    var sheet = ss.getSheetByName(TAB_NAME);
    if (!sheet) return { ok: false, error: 'Tab "' + TAB_NAME + '" not found.' };

    if (targetRow) {
      // Repeat contact — update only # of calls (col H) and last call date (col I).
      // First onboard date, name, org, etc. are left untouched since they're historical.
      sheet.getRange(targetRow, 8).setValue(row.numCalls || "");
      sheet.getRange(targetRow, 9).setValue(isoStringToDisplayDate(row.lastCallDate));
      return { ok: true, updatedRow: targetRow };
    }

    var values = [
      row.userName || "",
      row.userEmail || "",
      row.organization || "",
      row.userType || "",
      isoStringToDisplayDate(row.firstOnboardDate),
      row.salesperson || "",
      row.onboarder || "",
      row.numCalls || "",
      isoStringToDisplayDate(row.lastCallDate),
      row.referral || "",
      row.testimonial || ""
    ];

    sheet.insertRowBefore(2);
    sheet.getRange(2, 1, 1, values.length).setValues([values]);

    return { ok: true, insertedAt: 2 };
  } catch (e) {
    return { ok: false, error: e.message || String(e) };
  } finally {
    lock.releaseLock();
  }
}

// ── TEST ONLY — verifies the sheet write independent of HubSpot/Fathom ──
// Deliberately hardcoded to the TEST sheet ID, NOT the SHEET_ID constant
// above — this inserts a literal dummy row, and must never be able to
// write into the real production sheet even if SHEET_ID gets pointed
// there and someone runs this manually from the Apps Script editor.
function testSheetWrite() {
  var TEST_ONLY_SHEET_ID = "YOUR_TEST_TRACKING_SHEET_ID_HERE";
  var dummyRow = [
    "TEST USER", "test@example.com", "Test Org Inc.", "Individual",
    "6/30/26", "Test Salesperson", "Test Onboarder", "1", "6/30/26", "N", "N"
  ];
  var ss = SpreadsheetApp.openById(TEST_ONLY_SHEET_ID);
  var sheet = ss.getSheetByName(TAB_NAME);
  if (!sheet) { Logger.log('ERROR: Tab "' + TAB_NAME + '" not found.'); return; }
  sheet.insertRowBefore(2);
  sheet.getRange(2, 1, 1, dummyRow.length).setValues([dummyRow]);
  Logger.log("SUCCESS — dummy row inserted at row 2 of the TEST sheet only.");
}