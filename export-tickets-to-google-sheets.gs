function onOpen() {
  var ui = SpreadsheetApp.getUi()
    .createMenu("ZD stats")
    .addItem("Refresh tickets", "zd_tickets")
    .addToUi();
}

function search_array(my_array, field_id, field_value) {
  var r = {};
  if (typeof my_array == "object") {
    my_array.forEach(function (e, i) {
      if (field_id in e && e[field_id] == field_value) {
        r = e;
      }
    });
  }
  return r;
}

function zd_get_incremental_tickets() {
  var tickets = [];
  var users = [];
  var groups = [];
  var orgs = [];
  var url =
    "https://xxx.zendesk.com/api/v2/incremental/tickets.json?per_page=1000&start_time=1609459200&include=users,groups,organizations,metric_sets,ticket_forms"; // Replace XXX by your company ZD subdomain
  var end_of_stream = false;
  while (end_of_stream == false) {
    var response = UrlFetchApp.fetch(url, {
      headers: {
        Authorization:
          "Basic " +
          Utilities.base64Encode(
            "ZENDESK EMAIL/token:ZENDESK TOKEN"
          ),
      },
    });

    resp = JSON.parse(response.getContentText());
    url = resp["next_page"];
    end_of_stream = resp["end_of_stream"];
    resp["tickets"].forEach(function (e, i) {
      tickets.push(e);
    });
    resp["users"].forEach(function (e, i) {
      users.push(e);
    });
    resp["groups"].forEach(function (e, i) {
      groups.push(e);
    });
    resp["organizations"].forEach(function (e, i) {
      orgs.push(e);
    });
    var forms = resp["ticket_forms"];
  }
  return [tickets, users, groups, orgs, forms];
}

function zd_tickets() {
  // Clear sheet and set title
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tickets"); // Set the spreadsheet name as appropriate
  if (sheet != null) {
    sheet.getRange("A:S").clear(); // Adjust these ranges to the columns you want to display
    sheet
      .getRange("A1:S1")
      .setValues([
        [
          "Ticket ID",
          "Creation date",
          "Assign date",
          "Assign Time",
          "Due date",
          "Solve date",
          "Status",
          "Priority",
          "Type",
          "SOME_CUSTOM_FIELD1",
          "Requester Organization",
          "Group",
          "Assignee",
          "SOME_CUSTOM_FIELD2",
          "Subject",
          "SOME_CUSTOM_FIELD3",
          "Satisfaction Rating",
          "Full Resolution Time",
          "Ticket form",
        ],
      ]);
  }
  var row = 2;

  // Set variables
  var incrementals = zd_get_incremental_tickets();
  var tickets = incrementals[0];
  var users = incrementals[1];
  var groups = incrementals[2];
  var orgs = incrementals[3];
  var forms = incrementals[4];

  // Fill Support tickets sheet
  tickets.forEach(function (e, i) {
    if (e["status"] != "deleted") {
      var assign_time = "";
      if (e["metric_set"]["initially_assigned_at"]) {
        assign_time = Math.floor(
          (new Date(e["metric_set"]["initially_assigned_at"]).getTime() -
            new Date(e["created_at"]).getTime()) /
            (60 * 1000)
        );
      }
      var vals = [
        e["id"],
        e["created_at"],
        e["metric_set"]["initially_assigned_at"],
        assign_time,
        e["due_at"],
        e["metric_set"]["solved_at"],
        e["status"],
        e["priority"],
        e["type"],
        search_array(e["custom_fields"], "id", "CUSTOM_FIELD_1_ID")["value"],
        search_array(orgs, "id", e["organization_id"])["name"],
        search_array(groups, "id", e["group_id"])["name"],
        search_array(users, "id", e["assignee_id"])["name"],
        search_array(e["custom_fields"], "id", "CUSTOM_FIELD_2_ID")["value"],
        e["subject"],
        search_array(e["custom_fields"], "id", "CUSTOM_FIELD_3_ID")["value"],
        e["satisfaction_rating"]["score"],
        e["metric_set"]["full_resolution_time_in_minutes"]["business"],
        search_array(forms, "id", e["ticket_form_id"])["name"],
      ];
      sheet.getRange(row, 1, 1, 19).setValues([vals]);
      row += 1;
    }
  });

  // Create a filter
  if (sheet.getFilter() !== null) {
    sheet.getFilter().remove();
  }
  sheet.getRange(1, 1, row, 19).activate();
  sheet.getRange(1, 1, row, 19).createFilter();
  sheet.getRange("A1").activate();
  sheet.getFilter().sort(1, false);
}
