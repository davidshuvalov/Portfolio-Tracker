# Deployment Guide — Portfolio Tracker Commercial Infrastructure

## Architecture overview

```
Browser ──► Streamlit app (Streamlit Community Cloud or Render)
                │
                │  POST /api/create-checkout-session
                │  POST /api/create-billing-portal-session
                ▼
          FastAPI backend (Render)
                │
                ├──► Supabase (auth + profiles DB)
                └──► Stripe (subscriptions + webhooks)

Stripe ──► POST /api/stripe/webhook ──► FastAPI ──► Supabase profiles
```

---

## 1. Supabase setup

1. Create a project at [supabase.com](https://supabase.com).
2. Go to **SQL Editor** and run the contents of `supabase/schema.sql`.
3. In **Settings → Auth**, configure:
   - **Email confirmation**: Enable or disable depending on preference.
     If disabled, users are logged in immediately after signup (simpler UX).
   - **Site URL**: Set to your Streamlit app URL, e.g. `https://your-app.streamlit.app`.
   - **Redirect URLs**: Add `https://your-app.streamlit.app/**`.
4. Collect credentials from **Settings → API**:
   - **Project URL** (`SUPABASE_URL`)
   - **anon/public key** (`SUPABASE_ANON_KEY`) — used by the Streamlit frontend
   - **service_role key** (`SUPABASE_SERVICE_ROLE_KEY`) — used only by the FastAPI backend

---

## 2. Stripe setup

1. Create a [Stripe](https://stripe.com) account and activate it.
2. In the Stripe Dashboard create two **Products** with **recurring monthly pricing**:

   | Product | Price    | Nickname |
   |---------|----------|----------|
   | Lite    | $19 / mo | lite     |
   | Full    | $49 / mo | full     |

3. Copy each **Price ID** (`price_...`) — these become `STRIPE_PRICE_ID_LITE` and `STRIPE_PRICE_ID_FULL`.
4. In **Developers → API keys**, copy your **Secret key** (`STRIPE_SECRET_KEY`).
5. In **Developers → Webhooks**, add an endpoint:
   - URL: `https://your-backend.onrender.com/api/stripe/webhook`
   - Events to listen for:
     - `checkout.session.completed`
     - `customer.subscription.created`
     - `customer.subscription.updated`
     - `customer.subscription.deleted`
     - `invoice.paid`
     - `invoice.payment_failed`
   - Copy the **Signing secret** (`STRIPE_WEBHOOK_SECRET`).

---

## 3. FastAPI backend — deploy to Render

### Option A: Web Service (recommended)

1. Push this repo to GitHub.
2. Go to [render.com](https://render.com) → **New → Web Service**.
3. Connect your GitHub repo.
4. Configure:
   - **Root Directory**: `backend`
   - **Runtime**: Python 3
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `uvicorn main:app --host 0.0.0.0 --port $PORT`
5. Add environment variables (from `backend/.env.example`):

   ```
   SUPABASE_URL
   SUPABASE_SERVICE_ROLE_KEY
   STRIPE_SECRET_KEY
   STRIPE_WEBHOOK_SECRET
   STRIPE_PRICE_ID_LITE
   STRIPE_PRICE_ID_FULL
   FRONTEND_URL   ← your Streamlit app URL
   ```

6. Deploy. Note the service URL (e.g. `https://pt-backend.onrender.com`).

> **Free tier note**: Render free instances spin down after inactivity.
> Use a paid plan or add a keep-alive ping if this causes Stripe webhook timeouts.

### Option B: `render.yaml` (infrastructure-as-code)

Add a `render.yaml` at the repo root for one-click deploys:

```yaml
services:
  - type: web
    name: pt-backend
    runtime: python
    rootDir: backend
    buildCommand: pip install -r requirements.txt
    startCommand: uvicorn main:app --host 0.0.0.0 --port $PORT
    envVars:
      - key: SUPABASE_URL
        sync: false
      - key: SUPABASE_SERVICE_ROLE_KEY
        sync: false
      - key: STRIPE_SECRET_KEY
        sync: false
      - key: STRIPE_WEBHOOK_SECRET
        sync: false
      - key: STRIPE_PRICE_ID_LITE
        sync: false
      - key: STRIPE_PRICE_ID_FULL
        sync: false
      - key: FRONTEND_URL
        sync: false
```

---

## 4. Streamlit frontend — deploy to Streamlit Community Cloud

1. Push the repo to GitHub.
2. Go to [share.streamlit.io](https://share.streamlit.io) → **New app**.
3. Configure:
   - **Repository**: your GitHub repo
   - **Branch**: `main`
   - **Main file path**: `v2/app.py`
4. In **Advanced settings → Secrets**, paste the contents of
   `v2/.streamlit/secrets.toml.example` with real values filled in:

   ```toml
   SUPABASE_URL = "https://xxxxxxxxxxxxxxxxxxxx.supabase.co"
   SUPABASE_ANON_KEY = "eyJhbGc..."
   BACKEND_URL = "https://pt-backend.onrender.com"
   ```

5. Deploy. Note the app URL and set it as `FRONTEND_URL` in the backend env vars
   and as **Site URL** in Supabase Auth settings.

### Alternative: deploy to Render

1. Go to Render → **New → Web Service**.
2. Configure:
   - **Root Directory**: `v2`
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `streamlit run app.py --server.port $PORT --server.address 0.0.0.0`
3. Add the same three secrets as environment variables.

---

## 5. Local development

### Backend

```bash
cd backend
cp .env.example .env          # fill in real values
pip install -r requirements.txt
uvicorn main:app --reload --port 8000
```

### Streamlit frontend

```bash
cd v2
cp .streamlit/secrets.toml.example .streamlit/secrets.toml   # fill in values
pip install -r requirements.txt
streamlit run app.py
```

For local Stripe webhook testing, use the [Stripe CLI](https://stripe.com/docs/stripe-cli):

```bash
stripe listen --forward-to localhost:8000/api/stripe/webhook
```

This gives you a local `STRIPE_WEBHOOK_SECRET` to use in your backend `.env`.

---

## 6. Post-deploy checklist

- [ ] Supabase SQL schema applied and trigger verified (create a test user, check `profiles` row appears)
- [ ] Stripe products and prices created, Price IDs set in backend env
- [ ] Stripe webhook endpoint registered and pointing at deployed backend URL
- [ ] Backend `/health` returns `{"status": "ok"}`
- [ ] Streamlit `BACKEND_URL` secret points at deployed backend
- [ ] Supabase **Site URL** and **Redirect URLs** set to Streamlit app URL
- [ ] End-to-end test: sign up → subscribe (use Stripe test card `4242 4242 4242 4242`) → app unlocks → strategy limit enforced

---

## 7. Plan enforcement reference

| Plan  | Price  | `subscription_plan` | `strategy_limit` | Gated features                  |
|-------|--------|---------------------|------------------|---------------------------------|
| None  | —      | `null`              | `0`              | Everything                      |
| Lite  | $19/mo | `lite`              | `20`             | Full-plan analytics (see below) |
| Full  | $49/mo | `full`              | `null` (unlimited)| Nothing                         |

Full-plan-only features (gated via `gate_full_plan()` in `ui/plan_gate.py`):
- Portfolio Optimizer
- Leave One Out
- Margin Tracking
- Market Analysis
- Portfolio Compare
- Portfolio History

To gate a feature in any page:

```python
from ui.plan_gate import gate_full_plan

if not gate_full_plan("Portfolio Optimizer"):
    st.stop()
```

To enforce strategy count at the top of any page:

```python
from ui.plan_gate import enforce_strategy_limit

enforce_strategy_limit()
```
