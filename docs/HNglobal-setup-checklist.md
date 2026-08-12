# HNglobal Setup — Instructions & Checklist

**Goal:** One Cursor team named **HNglobal** with seats for `agivator@gmail.com` and `eyredn@gmail.com`, plus one GitHub org with private repos and a shared area.

> Print this page from your browser (Ctrl/Cmd+P), or download the PDF: [`HNglobal-setup-checklist.pdf`](./HNglobal-setup-checklist.pdf).

---

## Overview

| Layer | What to use | Who |
|---|---|---|
| **Cursor** | Upgrade existing account → Teams plan named **HNglobal** | Seat 1: `agivator@gmail.com` (Admin) · Seat 2: `eyredn@gmail.com` (Member or Admin) |
| **GitHub** | Create org **HNglobal** | Each person uses their own personal GitHub user and joins the org |
| **Code layout** | Repos under that org | Shared repo(s) both can access + private repo(s) limited by GitHub Teams |
| **Naming** | Brand/team = HNglobal | Personal emails stay as seat logins · Profile handles are permanent once claimed |

### Avoid

- Creating a second Cursor billing account just to rename
- Two separate Pro subscriptions
- Putting “private folders” inside one shared repo and expecting secrecy (GitHub permissions are **repo-level**, not folder-level)

---

## 1. Cursor — upgrade & name the team

Do this while signed in as **agivator@gmail.com**. One Cursor account can belong to only one team at a time.

1. Open [cursor.com/dashboard](https://cursor.com/dashboard) and sign in as agivator@gmail.com.
2. Choose **Upgrade to Teams** / **Create Team** (wording may vary).
3. Set the **team name** to **HNglobal**. Pick monthly or yearly billing.
4. Complete payment. Teams start on Standard seats; Premium (if available) can be upgraded per member later from the members menu.
5. Confirm you are a **paid Admin** on the team (needed if you use Cursor yourself).
6. Optional: set your display name in account settings. Optional profile handle at [cursor.com/profile](https://cursor.com/profile) — **cannot be changed** after claiming.

### Cursor checklist

- [ ] Signed in as agivator@gmail.com
- [ ] Upgraded / created Teams plan
- [ ] Team name set to **HNglobal**
- [ ] Billing payment completed
- [ ] agivator@gmail.com is Admin (paid seat)
- [ ] Display name updated (optional)
- [ ] Profile handle claimed only if sure (optional, permanent)

---

## 2. Cursor — invite the second seat

1. In the team dashboard, open **Members** / **Invite Members**.
2. Invite **eyredn@gmail.com**. Choose role: **Member** (typical) or **Admin** (if they should manage billing).
3. Have eyredn open the invite email, create/sign in to Cursor with that address, and **accept** the invite.
4. Confirm both seats show as active on the Members page (billing is per active paid seat, prorated).

### Second-seat checklist

- [ ] Invite sent to eyredn@gmail.com
- [ ] Role chosen (Member / Admin): _______________
- [ ] Invite accepted; eyredn can open Cursor
- [ ] Both seats visible as active on Members
- [ ] Passwords stored securely for both accounts (password manager)
- [ ] 2FA enabled on both accounts (recommended)

---

## 3. GitHub — org + personal accounts

Use a **GitHub Organization** for the brand. Each person keeps (or creates) a **personal** GitHub user, then joins the org. Do not put company work only on one person’s personal account.

1. Ensure each person has a personal GitHub account:
   - Person A (Agivator) — existing or new personal user
   - Person B (eyredn) — existing or new personal user (“separate GitHub account” for the second person)
2. Create organization **HNglobal** at [github.com/organizations/plan](https://github.com/organizations/plan) (Free org is fine for unlimited public/private repos).
3. Invite both personal GitHub users as org members. Make at least one an **Owner**.
4. Set org **Base permissions** to **No permission** (Settings → Member privileges) so people only see repos you explicitly grant.
5. Create GitHub Teams, for example:
   - **shared** — both people
   - **agivator-private** — Agivator only
   - **eyredn-private** — eyredn only

### GitHub account / org checklist

- [ ] Personal GitHub ready for Agivator: @_______________
- [ ] Personal GitHub ready for eyredn: @_______________
- [ ] Org **HNglobal** created
- [ ] Both users invited and accepted
- [ ] Owner role assigned: _______________
- [ ] Base permissions = No permission
- [ ] Team **shared** created; both members added
- [ ] Team **agivator-private** created
- [ ] Team **eyredn-private** created
- [ ] 2FA enabled on both GitHub users (recommended)

---

## 4. GitHub — private repos + shared area

Use **separate private repositories** for private work. A “shared directory” in practice = one or more shared repos both can access.

1. Create private repo **HNglobal/shared** (or several shared repos). Grant the **shared** team Write (or Maintain).
2. Create private repo(s) for Agivator-only work. Grant only **agivator-private**.
3. Create private repo(s) for eyredn-only work. Grant only **eyredn-private**.
4. Optional: move existing repos (e.g. from agivator-oss) into the HNglobal org via Settings → Transfer ownership; then re-apply team permissions.
5. Add a short README in each repo stating purpose and who owns it.

### Repo checklist

- [ ] **shared** repo(s) created · team access: Write
- [ ] Agivator private repo(s) created · access limited
- [ ] eyredn private repo(s) created · access limited
- [ ] Verified: eyredn cannot open Agivator-only repos (and vice versa)
- [ ] Verified: both can clone/push **shared**
- [ ] Existing repos transferred / permissions updated (if any)
- [ ] README added to each repo

---

## 5. Connect Cursor ↔ GitHub

1. On each person’s machine, sign into Cursor with their seat email.
2. Connect GitHub (Cursor settings / dashboard integrations) using that person’s *personal* GitHub account that is in the HNglobal org.
3. Grant the GitHub App / OAuth access to the **HNglobal** organization (org owner may need to approve the app).
4. Open a shared repo in Cursor and confirm Cloud Agents / pull requests work as expected.
5. Repeat a quick smoke test for a private repo on the account that should own it.

### Connect checklist

- [ ] Agivator: Cursor ↔ GitHub connected
- [ ] eyredn: Cursor ↔ GitHub connected
- [ ] HNglobal org approved access for the Cursor GitHub integration
- [ ] Shared repo opens and works in Cursor (both seats)
- [ ] Private-repo smoke test passed

---

## 6. Suggested order (one pass)

- [ ] ① Upgrade Cursor → name team HNglobal
- [ ] ② Invite eyredn@gmail.com and confirm seat active
- [ ] ③ Create GitHub org HNglobal + personal accounts / invites
- [ ] ④ Create Teams + shared / private repos + permissions
- [ ] ⑤ Connect each Cursor seat to GitHub and smoke-test
- [ ] ⑥ Store credentials in a password manager; enable 2FA everywhere

---

## Quick reference

| Item | Value |
|---|---|
| Cursor dashboard | https://cursor.com/dashboard |
| New Cursor team | https://cursor.com/team/new-team |
| Cursor profile / handle | https://cursor.com/profile |
| GitHub new org | https://github.com/organizations/plan |
| Seat 1 | agivator@gmail.com — Admin |
| Seat 2 | eyredn@gmail.com — Member / Admin |
| Brand name | HNglobal (Cursor team + GitHub org) |

---

## Sign-off

| | |
|---|---|
| Completed by | _______________________________ |
| Date | _______________ |
| Agivator initials | _____ |
| eyredn initials | _____ |

Keep passwords out of this printout; use a password manager.
