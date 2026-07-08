# Deploy on the Mac Mini (monthly auto-update)

Same convention as `vc-content-digest` (`docs/04-macmini-deploy.md`): headless Mac
Mini reached by SSH over Tailscale, a **dedicated standard user per project**,
scheduling by a **system LaunchDaemon** with a reverse-DNS name, secrets kept out
of the plist.

What this one does each month: scrape Compass → update the Excel → rebuild the HTML
widget → `git push`. GitHub Pages publishes the refreshed page automatically.

- Host: `konstantins-mac-mini` (`100.102.159.25`), local TZ Europe/Podgorica.
- Repo: `https://github.com/slogatskiy/agent-tally` (public; Pages from `/docs`).
- Service user: `comptally` (standard, like `vcdigest` / `bitrix`).
- Project path on the Mini: `/Users/comptally/projects/agent-tally`.

---

## A. Owner steps (Konstantin, SSH as `konstantin`, sudo)

Identical to the vcdigest playbook — only the name changes.

```bash
# A1 — create the standard service user (no -admin)
sudo sysadminctl -addUser comptally -fullName "Compass Tally Service" -password -

# A2 — let it use SSH (REQUIRED, else the key is ignored)
sudo dseditgroup -o edit -a comptally -t user com.apple.access_ssh
dseditgroup -o checkmember -m comptally com.apple.access_ssh     # expect "yes"

# A3 — install Stepan's public key
sudo -u comptally mkdir -p /Users/comptally/.ssh
sudo -u comptally tee /Users/comptally/.ssh/authorized_keys <<'KEY'
<paste Stepan's id_ed25519.pub>
KEY
sudo chmod 700 /Users/comptally/.ssh
sudo chmod 600 /Users/comptally/.ssh/authorized_keys
sudo chown -R comptally:staff /Users/comptally/.ssh

# A5 — install the LaunchDaemon (after B is done and a manual run works)
sudo cp /Users/comptally/projects/agent-tally/com.comptally.monthly.plist /Library/LaunchDaemons/
sudo chown root:wheel /Library/LaunchDaemons/com.comptally.monthly.plist
sudo chmod 644        /Library/LaunchDaemons/com.comptally.monthly.plist
sudo launchctl bootstrap system /Library/LaunchDaemons/com.comptally.monthly.plist
sudo launchctl kickstart -k system/com.comptally.monthly     # test run now
```

> Reusing the existing `vcdigest` user instead of making a new one also works if you
> want to skip A entirely — just clone into `/Users/vcdigest/projects/agent-tally`
> and set `UserName` in the plist to `vcdigest`. The dedicated user is only for
> clean separation, matching the house style.

## B. Dev steps (Stepan, SSH as `comptally`)

```bash
# B1 — clone
mkdir -p ~/projects && cd ~/projects
git clone https://github.com/slogatskiy/agent-tally.git
cd agent-tally && mkdir -p logs

# B2 — Python deps in a project venv (system python3 is fine; no uv needed)
python3 -m venv .venv
./.venv/bin/pip install -r requirements.txt

# B3 — git identity + push credentials (this is the only "secret")
git config user.name  "Compass Tally Bot"
git config user.email "comptally@users.noreply.github.com"
# Store a fine-grained PAT (repo: contents read/write on slogatskiy/agent-tally):
git config credential.helper store
git push            # prompts once for username=slogatskiy, password=<PAT>; cached after
#   — or, if gh is installed on the Mini:  gh auth login  (HTTPS, paste PAT)

# B4 — smoke test the whole pipeline by hand BEFORE arming the daemon
./run.sh
tail -n 40 logs/run.log        # should end with "published to GitHub Pages"
```

Then hand back to owner for step **A5**.

---

## Notes

- **If the Mini is asleep at 09:00 on the 1st**, launchd runs the missed job on
  wake, so a month is never skipped. Optional belt-and-suspenders:
  `sudo pmset repeat wakeorpoweron 1 8:55:00`.
- **Change cadence/time**: edit `StartCalendarInterval` in the plist, then
  `sudo launchctl bootout ...` and `bootstrap ...` again.
- **The PAT is the only credential.** Keep it fine-grained (single repo,
  contents:write). No SMTP / API keys here — this project only needs to scrape
  public pages and push to git.
- **Logs**: `logs/run.log` (the script) and `logs/daemon.{out,err}.log` (launchd).
- **Nothing to publish = no commit.** If Compass counts didn't change, `run.sh`
  exits without a commit, so the Pages history stays meaningful.
