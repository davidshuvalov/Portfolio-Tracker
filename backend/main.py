"""
Portfolio Tracker — FastAPI backend

Handles:
  - Stripe Checkout session creation
  - Stripe Billing Portal session creation
  - Stripe webhook processing

Run locally:
    uvicorn main:app --reload --port 8000

Environment:
    See .env.example for required variables.
"""

from __future__ import annotations

import os

from dotenv import load_dotenv

load_dotenv()

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from routers import checkout, webhooks

app = FastAPI(
    title="Portfolio Tracker API",
    description="Backend for Stripe + Supabase billing infrastructure",
    version="1.0.0",
)

# ── CORS ──────────────────────────────────────────────────────────────────────
# In production, restrict this to your Streamlit app URL.
frontend_url = os.getenv("FRONTEND_URL", "*")
app.add_middleware(
    CORSMiddleware,
    allow_origins=[frontend_url] if frontend_url != "*" else ["*"],
    allow_credentials=True,
    allow_methods=["POST", "GET", "OPTIONS"],
    allow_headers=["*"],
)

# ── Routers ───────────────────────────────────────────────────────────────────
app.include_router(checkout.router, prefix="/api", tags=["billing"])
app.include_router(webhooks.router, prefix="/api", tags=["webhooks"])


# ── Health check ──────────────────────────────────────────────────────────────
@app.get("/health", tags=["ops"])
def health():
    return {"status": "ok"}
