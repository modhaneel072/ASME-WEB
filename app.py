from datetime import date, timedelta
from flask import Flask, request, redirect, url_for, render_template_string, send_file
import pandas as pd

from models import db, Member, Item, Transaction

app = Flask(__name__)
app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///inventory.db"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
db.init_app(app)

# --- Helpers ---
def default_due_date(days=7):
    return date.today() + timedelta(days=days)

def parse_int(value, default=1):
    try:
        x = int(value)
        return x if x > 0 else default
    except Exception:
        return default

# --- Minimal inline HTML (so you can run without making templates yet) ---
PAGE = """

<header class="topbar">
  <div class="brand">
    <img src="/static/asme_logo.png" alt="ASME" class="logo" />
    <div>
      <div class="title">ASME Inventory</div>
      <div class="subtitle">University of Iowa · Checkout / Return</div>
    </div>
  </div>
  <nav class="nav">
    <a class="navlink" href="/">Home</a>
    <a class="navlink" href="/pair/member">Pair Member Tag</a>
    <a class="navlink" href="/pair/item">Pair Item Tag</a>
    <a class="navlink" href="/export">Export</a>
  </nav>
</header>

<main class="container">

<p>
  <a href="/pair/member">Pair Member Tag</a> |
  <a href="/pair/item">Pair Item Tag</a> |
  <a href="/export">Export to Excel</a>
</p>

<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <title>ASME Inventory</title>
  <style>
  :root{
    --iowa-gold:#FFCD00;
    --iowa-black:#000000;
    --ink:#111;
    --card:#ffffff;
    --bg:#f6f6f6;
    --border:#e7e7e7;
  }

  body{
    font-family: system-ui, -apple-system, Arial;
    margin:0;
    background: var(--bg);
    color: var(--ink);
  }

  .topbar{
    background: var(--iowa-black);
    color: white;
    border-bottom: 4px solid var(--iowa-gold);
    padding: 16px 22px;
    display:flex;
    align-items:center;
    justify-content:space-between;
    gap:16px;
  }

  .brand{
    display:flex;
    align-items:center;
    gap:14px;
    min-width: 320px;
  }

  .logo{
    height: 44px;
    width: auto;
    display:block;
  }

  .title{
    font-size: 22px;
    font-weight: 800;
    letter-spacing: 0.2px;
    line-height: 1.1;
  }

  .subtitle{
    font-size: 12.5px;
    color: rgba(255,255,255,0.75);
    margin-top: 3px;
  }

  .nav{
    display:flex;
    gap:10px;
    flex-wrap:wrap;
    justify-content:flex-end;
  }

  .navlink{
    color: white;
    text-decoration:none;
    padding: 8px 10px;
    border-radius: 10px;
    border: 1px solid rgba(255,255,255,0.18);
    font-size: 13px;
  }

  .navlink:hover{
    border-color: var(--iowa-gold);
    box-shadow: 0 0 0 2px rgba(255,205,0,0.25) inset;
  }

  .container{
    max-width: 980px;
    margin: 22px auto;
    padding: 0 18px 40px;
  }

  .row { display:flex; gap:12px; flex-wrap:wrap; }
  .card{
    background: var(--card);
    border: 1px solid var(--border);
    padding: 16px;
    border-radius: 16px;
    margin: 14px 0;
    box-shadow: 0 6px 18px rgba(0,0,0,0.04);
  }

  input, select{
    padding: 10px;
    min-width: 240px;
    border-radius: 12px;
    border: 1px solid var(--border);
    outline: none;
  }

  input:focus, select:focus{
    border-color: var(--iowa-gold);
    box-shadow: 0 0 0 3px rgba(255,205,0,0.25);
  }

  button{
    padding: 10px 14px;
    border-radius: 12px;
    border: 1px solid #222;
    background: var(--iowa-black);
    color: white;
    cursor: pointer;
    font-weight: 650;
  }

  button:hover{
    border-color: var(--iowa-gold);
    box-shadow: 0 0 0 3px rgba(255,205,0,0.18);
  }

  a { color: #0b57d0; }
  .ok{ color:#0a7; font-weight:650; }
  .err{ color:#c00; font-weight:650; }
  .muted { color:#666; }

  table{ width:100%; border-collapse:collapse; margin-top:8px; }
  th, td{ border-bottom:1px solid #eee; padding: 10px 8px; text-align:left; }
  th{ font-size: 13px; color:#333; }
</style>
</head>
<body
  <h1>ASME Inventory</h1>
  <p class="muted">Workflow: scan/enter Member → scan/enter Item → checkout/return.</p>

  {% if message %}
    <p class="{{ 'ok' if success else 'err' }}">{{ message }}</p>
  {% endif %}

  <div class="card">
    <h2>Checkout / Return</h2>
    <form method="POST" action="/transact">
      <div class="row">
        <div>
          <label><b>Member (scan NFC tag or pick)</b></label><br/>
          <input name="member_tag" placeholder="Scan member tag here (optional)" autofocus />
          <div class="muted">If you don't use member NFC tags, just select from dropdown.</div>
          <br/>
          <select name="member_id">
            <option value="">-- or select member --</option>
            {% for m in members %}
              <option value="{{m.id}}">{{m.name}} ({{m.email}}) - {{m.member_class}}</option>
            {% endfor %}
          </select>
        </div>

        <div>
          <label><b>Item (scan NFC tag or pick)</b></label><br/>
          <input name="item_tag" placeholder="Scan item tag here (optional)" />
          <div class="muted">If you tag bins/items, scan here for instant lookup.</div>
          <br/>
          <select name="item_id">
            <option value="">-- or select item --</option>
            {% for it in items %}
              <option value="{{it.id}}">{{it.name}} (avail {{it.available_qty}} / {{it.total_qty}})</option>
            {% endfor %}
          </select>
        </div>
      </div>

      <div class="row" style="margin-top:12px;">
        <div>
          <label><b>Action</b></label><br/>
          <select name="action">
            <option value="checkout">Checkout</option>
            <option value="return">Return</option>
          </select>
        </div>

        <div>
          <label><b>Qty</b></label><br/>
          <input name="qty" value="1" />
        </div>

        <div>
          <label><b>Due date</b> (optional)</label><br/>
          <input name="due_date" value="{{default_due}}" placeholder="YYYY-MM-DD" />
        </div>
      </div>

      <div style="margin-top:12px;">
        <label><b>Notes</b> (optional)</label><br/>
        <input name="notes" placeholder="e.g., 'DBF team use'" style="min-width: 600px;" />
      </div>

      <div style="margin-top:14px;">
        <button type="submit">Submit</button>
        <a href="/export" style="margin-left: 10px;">Export to Excel</a>
      </div>
    </form>
  </div>

  <div class="card">
    <h2>Recent Transactions</h2>
    <table>
      <thead>
        <tr>
          <th>Time</th><th>Member</th><th>Item</th><th>Action</th><th>Qty</th><th>Due</th>
        </tr>
      </thead>
      <tbody>
        {% for t in recent %}
        <tr>
          <td>{{t.timestamp.strftime("%Y-%m-%d %H:%M")}}</td>
          <td>{{t.member.name}}</td>
          <td>{{t.item.name}}</td>
          <td>{{t.action}}</td>
          <td>{{t.qty}}</td>
          <td>{{t.due_date if t.due_date else ""}}</td>
        </tr>
        {% endfor %}
      </tbody>
    </table>
  </div>

</body>
</html>
"""

@app.get("/")
def index():
    members = Member.query.order_by(Member.name.asc()).all()
    items = Item.query.order_by(Item.name.asc()).all()
    recent = Transaction.query.order_by(Transaction.timestamp.desc()).limit(15).all()
    return render_template_string(
        PAGE,
        members=members,
        items=items,
        recent=recent,
        default_due=str(default_due_date()),
        message=None,
        success=True
    )

@app.post("/transact")
def transact():
    member_tag = (request.form.get("member_tag") or "").strip() or None
    item_tag = (request.form.get("item_tag") or "").strip() or None

    member_id = request.form.get("member_id") or None
    item_id = request.form.get("item_id") or None

    action = (request.form.get("action") or "").strip().lower()
    qty = parse_int(request.form.get("qty"), default=1)
    notes = (request.form.get("notes") or "").strip() or None

    due_raw = (request.form.get("due_date") or "").strip()
    due = None
    if due_raw:
        try:
            y, m, d = [int(x) for x in due_raw.split("-")]
            due = date(y, m, d)
        except Exception:
            due = None  # ignore bad format

    # Find member
    member = None
    if member_tag:
        member = Member.query.filter_by(nfc_tag=member_tag).first()
    if not member and member_id:
        member = Member.query.get(int(member_id))

    # Find item
    item = None
    if item_tag:
        item = Item.query.filter_by(nfc_tag=item_tag).first()
    if not item and item_id:
        item = Item.query.get(int(item_id))

    # Validate
    if not member:
        return _render_with_msg("Could not find member. Scan a member tag or pick a member.", success=False)
    if not item:
        return _render_with_msg("Could not find item. Scan an item tag or pick an item.", success=False)
    if action not in {"checkout", "return"}:
        return _render_with_msg("Invalid action.", success=False)

    # Apply inventory rules
    if action == "checkout":
        if item.available_qty < qty:
            return _render_with_msg(
                f"Not enough available. {item.name} has {item.available_qty} available.",
                success=False
            )
        item.available_qty -= qty
    else:  # return
        # Don’t exceed total_qty (keeps data sane)
        item.available_qty = min(item.total_qty, item.available_qty + qty)

    # Log transaction
    t = Transaction(
        member_id=member.id,
        item_id=item.id,
        action=action,
        qty=qty,
        due_date=due if action == "checkout" else None,
        notes=notes
    )
    db.session.add(t)
    db.session.commit()

    return redirect(url_for("index"))

def _render_with_msg(message, success):
    members = Member.query.order_by(Member.name.asc()).all()
    items = Item.query.order_by(Item.name.asc()).all()
    recent = Transaction.query.order_by(Transaction.timestamp.desc()).limit(15).all()
    return render_template_string(
        PAGE,
        members=members,
        items=items,
        recent=recent,
        default_due=str(default_due_date()),
        message=message,
        success=success
    )

@app.get("/export")
def export_excel():
    # Export members, items, and transactions to one Excel file with 3 sheets
    members = Member.query.all()
    items = Item.query.all()
    tx = Transaction.query.order_by(Transaction.timestamp.desc()).all()

    members_df = pd.DataFrame([{
        "id": m.id, "name": m.name, "email": m.email, "class": m.member_class, "nfc_tag": m.nfc_tag, "created_at": m.created_at
    } for m in members])

    items_df = pd.DataFrame([{
        "id": i.id, "name": i.name, "category": i.category, "location": i.location,
        "total_qty": i.total_qty, "available_qty": i.available_qty, "nfc_tag": i.nfc_tag, "created_at": i.created_at
    } for i in items])

    tx_df = pd.DataFrame([{
        "id": t.id,
        "timestamp": t.timestamp,
        "member": t.member.name,
        "email": t.member.email,
        "class": t.member.member_class,
        "item": t.item.name,
        "action": t.action,
        "qty": t.qty,
        "due_date": t.due_date,
        "notes": t.notes
    } for t in tx])

    filename = "inventory_export.xlsx"
    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        members_df.to_excel(writer, index=False, sheet_name="Members")
        items_df.to_excel(writer, index=False, sheet_name="Items")
        tx_df.to_excel(writer, index=False, sheet_name="Transactions")

    return send_file(filename, as_attachment=True)

if __name__ == "__main__":
    with app.app_context():
        db.create_all()
    app.run(debug=True)

# ----------------------------
# NFC Pairing (Admin-lite)
# ----------------------------

PAIR_MEMBER_PAGE = """
<!doctype html>
<html>
<head><meta charset="utf-8"/><title>Pair Member Tag</title></head>
<body style="font-family:system-ui; margin:24px; max-width:900px;">
  <h1>Pair Member NFC Tag</h1>
  <p><a href="/">← Back to Home</a></p>

  {% if message %}
    <p style="color: {{ 'green' if success else 'crimson' }};">{{ message }}</p>
  {% endif %}

  <form method="POST">
    <label><b>Select member</b></label><br/>
    <select name="member_id" required style="padding:10px; min-width:420px;">
      <option value="">-- choose --</option>
      {% for m in members %}
        <option value="{{m.id}}">{{m.name}} ({{m.email}}) - {{m.member_class}}</option>
      {% endfor %}
    </select>
    <br/><br/>

    <label><b>Scan member NFC tag</b> (will “type” an ID)</label><br/>
    <input name="tag" autofocus required placeholder="Tap tag/card here"
           style="padding:10px; min-width:420px;" />
    <br/><br/>

    <button type="submit" style="padding:10px 14px;">Pair Tag</button>
  </form>

  <h3 style="margin-top:28px;">Currently Paired Members</h3>
  <table style="border-collapse:collapse; width:100%;">
    <tr><th align="left">Member</th><th align="left">Tag</th></tr>
    {% for m in paired %}
      <tr style="border-top:1px solid #eee;">
        <td style="padding:8px 0;">{{m.name}} ({{m.email}})</td>
        <td style="padding:8px 0;">{{m.nfc_tag}}</td>
      </tr>
    {% endfor %}
  </table>
</body>
</html>
"""

@app.get("/pair/member")
def pair_member_get():
    members = Member.query.order_by(Member.name.asc()).all()
    paired = Member.query.filter(Member.nfc_tag.isnot(None)).order_by(Member.name.asc()).all()
    return render_template_string(PAIR_MEMBER_PAGE, members=members, paired=paired, message=None, success=True)

@app.post("/pair/member")
def pair_member_post():
    member_id = int(request.form["member_id"])
    tag = request.form["tag"].strip()

    # Prevent duplicate tag assignment
    if Member.query.filter_by(nfc_tag=tag).first():
        return _pair_member_msg("That tag is already assigned to a member.", success=False)

    m = Member.query.get(member_id)
    if not m:
        return _pair_member_msg("Member not found.", success=False)

    m.nfc_tag = tag
    db.session.commit()
    return _pair_member_msg(f"✅ Paired tag to {m.name}.", success=True)

def _pair_member_msg(message, success):
    members = Member.query.order_by(Member.name.asc()).all()
    paired = Member.query.filter(Member.nfc_tag.isnot(None)).order_by(Member.name.asc()).all()
    return render_template_string(PAIR_MEMBER_PAGE, members=members, paired=paired, message=message, success=success)


PAIR_ITEM_PAGE = """
<!doctype html>
<html>
<head><meta charset="utf-8"/><title>Pair Item Tag</title></head>
<body style="font-family:system-ui; margin:24px; max-width:900px;">
  <h1>Pair Item NFC Tag</h1>
  <p><a href="/">← Back to Home</a></p>

  {% if message %}
    <p style="color: {{ 'green' if success else 'crimson' }};">{{ message }}</p>
  {% endif %}

  <form method="POST">
    <label><b>Select item</b></label><br/>
    <select name="item_id" required style="padding:10px; min-width:420px;">
      <option value="">-- choose --</option>
      {% for it in items %}
        <option value="{{it.id}}">{{it.name}} (avail {{it.available_qty}} / {{it.total_qty}})</option>
      {% endfor %}
    </select>
    <br/><br/>

    <label><b>Scan item NFC tag</b></label><br/>
    <input name="tag" autofocus required placeholder="Tap tag on bin/item here"
           style="padding:10px; min-width:420px;" />
    <br/><br/>

    <button type="submit" style="padding:10px 14px;">Pair Tag</button>
  </form>

  <h3 style="margin-top:28px;">Currently Paired Items</h3>
  <table style="border-collapse:collapse; width:100%;">
    <tr><th align="left">Item</th><th align="left">Tag</th></tr>
    {% for it in paired %}
      <tr style="border-top:1px solid #eee;">
        <td style="padding:8px 0;">{{it.name}}</td>
        <td style="padding:8px 0;">{{it.nfc_tag}}</td>
      </tr>
    {% endfor %}
  </table>
</body>
</html>
"""

@app.get("/pair/item")
def pair_item_get():
    items = Item.query.order_by(Item.name.asc()).all()
    paired = Item.query.filter(Item.nfc_tag.isnot(None)).order_by(Item.name.asc()).all()
    return render_template_string(PAIR_ITEM_PAGE, items=items, paired=paired, message=None, success=True)

@app.post("/pair/item")
def pair_item_post():
    item_id = int(request.form["item_id"])
    tag = request.form["tag"].strip()

    # Prevent duplicate tag assignment
    if Item.query.filter_by(nfc_tag=tag).first():
        return _pair_item_msg("That tag is already assigned to an item.", success=False)

    it = Item.query.get(item_id)
    if not it:
        return _pair_item_msg("Item not found.", success=False)

    it.nfc_tag = tag
    db.session.commit()
    return _pair_item_msg(f"✅ Paired tag to {it.name}.", success=True)

def _pair_item_msg(message, success):
    items = Item.query.order_by(Item.name.asc()).all()
    paired = Item.query.filter(Item.nfc_tag.isnot(None)).order_by(Item.name.asc()).all()
    return render_template_string(PAIR_ITEM_PAGE, items=items, paired=paired, message=message, success=success)
