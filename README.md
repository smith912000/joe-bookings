# Joe's bookings

Two pages, no server needed (works on GitHub Pages):

- `index.html` — customers: week calendar, green = free, grey = not available. Tapping a free time goes to the existing sign-up form. Shows no client names.
- `admin.html` — Joe: load bookings CSV, add/move/cancel, save the CSV and the availability file. Moving a booking records the old time in "Changed from".

Source of truth is the Google Sheet "Joe's Bookings" (tab Bookings). `tools/sheet_to_json.py` turns its CSV export into `data/busy.json`, which the customer page reads. Weekly hours and gyms live in `data/config.json` (from the earlier My Week settings). Put the logo at `assets/logo.png`.

## Schedule (data/config.json)
- `open`: windows clients can book. `blocked`: Pure Gym shifts and Boxing Academy classes. `away`: whole-day absences. Joe edits all of these on `admin.html` and saves a new `config.json`.
- Assumptions Joe can change: Monday evening open 17:00-21:00; Friday open 10:00-17:00 (after his Pure Gym shift); nothing else is open unless added; Sunday 8-12 is a shift; class slots are one hour each.
- Flexible clients: sessions are lined up back to back with other PTs that day (booking note shown on the customer page).
- Changing a time: Joe's page keeps the client's details when a booking is moved (old time kept in "Changed from"); on the customer page a client's name and contact are remembered on their device and a change request needs only the new time.
