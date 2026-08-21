// ReneCode.gs — Rene's own backend functions, kept in a separate file
// from Lupe's Code.gs purely for organizational clarity — Apps Script
// has no real per-file isolation, all .gs files in one project share one
// single global namespace. Every function here is prefixed rene_
// specifically to guard against ever silently colliding with one of
// Lupe's own existing function names, which would otherwise happen
// silently with no error at all.
//
// Deliberately reuses Lupe's own hsFetch() and getProp() helpers rather
// than redefining equivalent versions — same shared HubSpot token, same
// already-tested request/error-handling behavior. Both of those live in
// Code.gs and are available here automatically, since this file shares
// the same project.

// The dedicated pipeline HubSpot's Deal Stage properties confirm as
// "OSBO Pipeline" — the company-wide bulk onboarding deal always lives
// here, separate from the "OSP/DV Pipeline" (property value "default")
// used for individual one-off broker deals. RENE is scoped to OSBO only:
// it should never pull from, or count, deals sitting in OSP/DV Pipeline.
var RENE_OSBO_PIPELINE_ID = "759448220";

// Extracts a HubSpot company ID from a pasted URL or raw ID — mirrors
// Lupe's own extractContactId(), just matching company-specific URL
// patterns (object type 0-2, or the older /company/ path) instead of
// contact ones.
function rene_extractCompanyId(input) {
  if (!input) return null;
  var s = String(input).trim();
  var patterns = [
    /\/company\/(\d+)/,
    /\/record\/0-2\/(\d+)/,
    /companyId=(\d+)/
  ];
  for (var i = 0; i < patterns.length; i++) {
    var m = s.match(patterns[i]);
    if (m) return m[1];
  }
  if (/^\d+$/.test(s)) return s; // plain ID
  return null;
}

// Main entry point called from ReneIndex.html. Takes a pasted company
// URL/ID, finds that company's OSBO-pipeline deal (RENE only ever cares
// about OSBO — individual OSP/DV broker deals on the same company are
// deliberately excluded, not aggregated in), and returns every contact
// on that deal, grouped by their team_role value.
//
// This first version deliberately only PULLS and DISPLAYS data — it does
// not write anything to the tracking sheet or enroll anyone in anything.
function rene_pullCompanyContacts(companyUrlOrId) {
  var companyId = rene_extractCompanyId(companyUrlOrId);
  if (!companyId) {
    return { ok: false, error: "Couldn't find a company ID in that URL. Paste the HubSpot company record URL, or a plain company ID." };
  }

  try {
    // Step 1: confirm the company exists, get its real name.
    var companyData = hsFetch("/crm/v3/objects/companies/" + companyId + "?properties=name", "get");
    var companyName = (companyData.properties && companyData.properties.name) || "(unnamed company)";

    // Step 2: get every deal associated with this company (just IDs).
    var dealAssoc = hsFetch("/crm/v4/objects/companies/" + companyId + "/associations/deals", "get");
    var dealIds = ((dealAssoc && dealAssoc.results) || []).map(function(r) { return r.toObjectId; });

    if (!dealIds.length) {
      return { ok: true, companyName: companyName, totalDeals: 0, totalContacts: 0, groups: rene_emptyGroups(),
               warning: "This company has no associated deals — nothing to pull contacts from." };
    }

    // Step 3: batch-fetch dealname + pipeline for every deal on the
    // company, so we can filter down to ONLY the OSBO-pipeline deal(s).
    // Individual OSP/DV broker deals live on the same company record but
    // are never what RENE should be pulling contacts from.
    var dealDetails = [];
    for (var i = 0; i < dealIds.length; i += 100) {
      var chunk = dealIds.slice(i, i + 100);
      var batch = hsFetch("/crm/v3/objects/deals/batch/read", "post", {
        inputs: chunk.map(function(id) { return { id: id }; }),
        properties: ["dealname", "pipeline"]
      });
      ((batch && batch.results) || []).forEach(function(d) { dealDetails.push(d); });
    }

    var osboDeals = dealDetails.filter(function(d) {
      return d.properties && d.properties.pipeline === RENE_OSBO_PIPELINE_ID;
    });

    if (!osboDeals.length) {
      return { ok: true, companyName: companyName, totalDeals: 0, totalContacts: 0, groups: rene_emptyGroups(),
               warning: "No OSBO Pipeline deal found for this company (" + dealIds.length +
                        " other deal(s) exist, but none in the OSBO Pipeline) — nothing to pull." };
    }

    // Exactly one OSBO deal is the expected, normal case — auto-select it.
    if (osboDeals.length === 1) {
      var only = osboDeals[0];
      var result = rene_pullContactsForDeal(only.id);
      result.companyName = companyName;
      result.dealName = (only.properties && only.properties.dealname) || "(unnamed deal)";
      result.totalDeals = 1;
      return result;
    }

    // More than one OSBO-pipeline deal on this company (rare — e.g. an
    // original deal plus a renewal deal). Rather than guessing which one
    // is correct or silently merging both, surface both as options and
    // let the person pick — matches the human-in-the-loop preference
    // used everywhere else in this project.
    return {
      ok: true,
      companyName: companyName,
      needsDealSelection: true,
      deals: osboDeals.map(function(d) {
        return { id: d.id, name: (d.properties && d.properties.dealname) || "(unnamed deal)" };
      })
    };

  } catch (e) {
    return { ok: false, error: e.message || String(e) };
  }
}

// Called from ReneIndex.html only when rene_pullCompanyContacts came back
// with needsDealSelection: true, after the person picks one of the listed
// OSBO deals manually. companyName/dealName are passed straight through
// from that earlier response purely for display — no extra lookup needed.
function rene_pullDealContacts(dealId, companyName, dealName) {
  try {
    var result = rene_pullContactsForDeal(dealId);
    result.companyName = companyName;
    result.dealName = dealName;
    result.totalDeals = 1;
    return result;
  } catch (e) {
    return { ok: false, error: e.message || String(e) };
  }
}

// Pulls every contact associated with ONE specific deal, batch-fetches
// their real name/email/team_role, and groups them. Shared by both the
// auto-detect path above and the manual-selection path, since the actual
// contact-pulling-and-grouping work is identical either way — only how
// the target deal gets chosen differs.
function rene_pullContactsForDeal(dealId) {
  var contactAssoc = hsFetch("/crm/v4/objects/deals/" + dealId + "/associations/contacts", "get");
  var contactIds = ((contactAssoc && contactAssoc.results) || []).map(function(r) { return r.toObjectId; });

  if (!contactIds.length) {
    return { ok: true, dealId: dealId, totalContacts: 0, groups: rene_emptyGroups(),
             warning: "This deal has no associated contacts." };
  }

  // Batch-fetch full details for every contact found, in chunks of 100
  // (HubSpot's own batch-read limit) — not one request per contact.
  var allContacts = [];
  for (var i = 0; i < contactIds.length; i += 100) {
    var chunk = contactIds.slice(i, i + 100);
    var batch = hsFetch("/crm/v3/objects/contacts/batch/read", "post", {
      inputs: chunk.map(function(id) { return { id: id }; }),
      properties: ["firstname", "lastname", "email", "team_role"]
    });
    ((batch && batch.results) || []).forEach(function(c) { allContacts.push(c); });
  }

  // Group by team_role. Confirmed directly against real portal data that
  // multi-value team_role comes back as ONE string with semicolons
  // separating each selected value (HubSpot's own multi-checkbox
  // format) — e.g. "Upsell Primary Contact;Primary Contact;Broker" — not
  // an array. A contact carrying more than one relevant role (e.g. both
  // Broker and Partner/Owner) genuinely appears in both of those groups
  // here, since that reflects their real, actual tagging rather than
  // forcing them into a single bucket.
  var groups = rene_emptyGroups();
  allContacts.forEach(function(c) {
    var p = c.properties || {};
    var firstName = p.firstname || "";
    var lastName = p.lastname || "";
    var email = p.email || "";
    var fullName = (firstName + " " + lastName).trim() || email || "(no name)";
    var roleRaw = p.team_role || "";
    var roles = roleRaw.split(";").map(function(r) { return r.trim(); }).filter(function(r) { return r; });

    var contactObj = { id: c.id, name: fullName, email: email, roles: roles };

    if (!roles.length) {
      groups.unknown.push(contactObj);
      return;
    }
    var matchedAny = false;
    roles.forEach(function(r) {
      var rl = r.toLowerCase();
      if (rl.indexOf("broker") !== -1) { groups.broker.push(contactObj); matchedAny = true; }
      if (rl.indexOf("admin") !== -1 || rl.indexOf("transaction coordinator") !== -1) { groups.admin.push(contactObj); matchedAny = true; }
      if (rl.indexOf("partner") !== -1 || rl.indexOf("owner") !== -1) { groups.partner.push(contactObj); matchedAny = true; }
    });
    // A contact with a team_role value that doesn't match any of the
    // three known categories still needs somewhere to land, rather than
    // silently vanishing from the total count.
    if (!matchedAny) groups.unknown.push(contactObj);
  });

  return { ok: true, dealId: dealId, totalContacts: allContacts.length, groups: groups };
}

function rene_emptyGroups() {
  return { broker: [], admin: [], partner: [], unknown: [] };
}

// Searches HubSpot companies by name — same conditions as Lupe's own
// searchContactsByName, adapted for companies. Genuinely simpler here:
// companies have one single "name" property, not a firstname/lastname
// split, so none of the multi-word cross-product logic Lupe needs (to
// handle a name being split across two separate fields) applies — a
// single wildcard-wrapped CONTAINS_TOKEN against name covers a full
// multi-word company name in one filter, the same way Lupe already
// handles a single-word query or an email.
function rene_searchCompaniesByName(query) {
  if (!query) return [];
  var trimmed = query.trim();
  if (!trimmed) return [];

  try {
    var data = hsFetch("/crm/v3/objects/companies/search", "post", {
      filterGroups: [
        { filters: [{ propertyName: "name", operator: "CONTAINS_TOKEN", value: "*" + trimmed + "*" }] }
      ],
      properties: ["name", "domain", "city", "state"],
      limit: 10
    });

    var results = data.results || [];
    return results.map(function(c) {
      var p = c.properties || {};
      return {
        id: c.id,
        name: p.name || "(no name)",
        domain: p.domain || "",
        location: [p.city, p.state].filter(function(x) { return x; }).join(", ")
      };
    });
  } catch (e) {
    return []; // fail quietly — same as Lupe's version, this just means the dropdown shows no results
  }
}