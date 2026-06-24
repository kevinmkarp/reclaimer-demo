import html
import json
import os
import re
from datetime import date
from io import BytesIO
from typing import Any

import streamlit as st
from dotenv import load_dotenv
from openai import OpenAI


try:
    from docx import Document
except Exception:  # pragma: no cover - optional dependency guard
    Document = None

try:
    from pypdf import PdfReader
except Exception:  # pragma: no cover - optional dependency guard
    PdfReader = None


APP_TITLE = "RE•CLAIMER"
ANALYSIS_VERSION = 3
MODEL = os.getenv("OPENAI_MODEL", "gpt-4.1")
DISCOVERY_MODEL = os.getenv("OPENAI_DISCOVERY_MODEL", "gpt-4.1-mini")
MAX_DISCOVERY_SOURCES = 6
MAX_DISCOVERY_SUMMARY_CHARS = 500
LIVE_WEB_DISABLED = os.getenv("RECLAIMER_DISABLE_LIVE_WEB", "0") == "1"
RESEARCH_MODES = (
    ["No web", "Evidence-only"]
    if LIVE_WEB_DISABLED
    else ["No web", "Web-grounded", "Evidence-only", "Evidence + web"]
)

SAMPLE_CLAIMS = {
    "Choose an example": {},
    "HelloFresh offer": {
        "brand": "HelloFresh",
        "claim": "Get 40% off HelloFresh for a year!",
        "category": "Meal kits",
    },
    "HelloFresh structured offer": {
        "brand": "HelloFresh",
        "claim": (
            "HelloFresh now has a 40% off for 52 weeks offer. Under a 2-person, "
            "3-meal plan, that brings average price to $6.99 per serving, the same "
            "as with EveryPlate, and less than Dinnerly at $7.99, shipping excluded. "
            "Plus, HelloFresh still has broader menu variety."
        ),
        "category": "Meal kits",
    },
    "Legal service fee": {
        "brand": "Sweet James Accident Lawyers",
        "claim": "No fee unless we win.",
        "category": "Personal injury law firm",
    },
    "Wearable size": {
        "brand": "Oura Ring",
        "claim": "40% smaller than Oura Ring 4.",
        "category": "Wearable health rings",
    },
    "Rent the Runway vague": {
        "brand": "Rent the Runway",
        "claim": "Rent the Runway is the best service for new outfits every week.",
        "category": "Clothing rental subscriptions",
    },
    "Rent the Runway structured": {
        "brand": "Rent the Runway",
        "claim": (
            "Rent the Runway offers up to 20 items per month, shipped in up to "
            "4 shipments of 5 items each, compared to Nuuly's 6 items per month "
            "in a single shipment."
        ),
        "category": "Clothing rental subscriptions",
    },
    "Running shoe durability": {
        "brand": "Roadster 2.0",
        "claim": "The most durable running shoe under $150.",
        "category": "Running shoes",
    },
    "SUV safety": {
        "brand": "",
        "claim": "The safest SUV for new parents.",
        "category": "SUVs",
    },
    "Payroll software": {
        "brand": "",
        "claim": "The easiest payroll software for small businesses.",
        "category": "Payroll software",
    },
    "Family hotel": {
        "brand": "",
        "claim": "The best family hotel in Nice.",
        "category": "Hotels",
    },
}

UNIVERSAL_RECOMMENDATION_QUESTION = (
    "Would this claim change how an AI recommendation agent answers consumer "
    "recommendation questions in this category?"
)

IMPACT_LEVELS = [
    "Ignored",
    "Brand context only",
    "Mention-worthy",
    "Consideration-worthy",
    "Strong consideration",
    "Conditional recommendation",
    "Recommendation-changing",
]

EVIDENCE_COVERAGE_LEVELS = [
    "Not checked",
    "Low",
    "Partial",
    "Moderate",
    "Strong",
    "Robust",
]


SYSTEM_PROMPT = """
You are an AI recommendation-agent evidence planner. Your job is to determine whether a marketing claim would be treated by an AI recommendation agent as decision evidence, brand context, or ignored, and what would need to change for it to affect a recommendation.

Do not assume the claim is true. Analyze whether the claim is specific, verifiable, comparative, current, caveated, and supported. Distinguish between claims an AI could safely repeat as brand marketing and claims it could state as fact. Identify what evidence, context, triangulation, and source corroboration would be needed before an AI agent should rely on the claim.

Infer claim type and evidence requirements from the claim itself. Do not use hardcoded category logic. If evidence or web context is provided, use it critically and cite source URLs when available. If web access is not enabled, do not imply that you checked the web. If evidence-only mode is selected, rely only on uploaded evidence and the user's inputs.

Do not stop at "needs substantiation." Show the path from vague marketing language to structured decision proof: measurable terms, comparison sets, caveats, source placement, and the internet surfaces where proof needs to live.

When numbers are present, perform a math and logic audit. Detect percentages, savings, price comparisons, "x times" claims, and percent smaller/larger claims. Check whether the wording accurately describes the math and identify what base, denominator, comparison, time period, and calculation method are missing.

Use this exact canonical impact ladder everywhere: Ignored, Brand context only, Mention-worthy, Consideration-worthy, Strong consideration, Conditional recommendation, Recommendation-changing. Use no custom variants. Do not show fake percentages.

Use this exact Evidence Coverage ladder and no other labels: Not checked, Low, Partial, Moderate, Strong, Robust.

Claim structure and evidence coverage are independent. A structurally strong claim can be Consideration-worthy even when Evidence Coverage is Not checked, Low, or Partial. Simulated Agent Impact should reflect source support; Claim as Written should reflect the payload's structure and decision relevance.

Impact definitions:
- Ignored: irrelevant, too vague, off-topic, contradictory, or unusable.
- Brand context only: describes positioning but does not affect recommendations.
- Mention-worthy: may be mentioned but does not enter the serious recommendation set by itself.
- Consideration-worthy: specific enough to enter the consideration set for a relevant need.
- Strong consideration: likely to be shortlisted or framed as a leading option.
- Conditional recommendation: likely recommended for a specific need with caveats.
- Recommendation-changing: strong enough to shift the likely default or preferred answer.

Evidence Coverage definitions:
- Not checked: no web queries or source scans.
- Low: one or two sources, mostly owned/targets, or unclear relevance.
- Partial: some relevant sources checked but major source types missing.
- Moderate: several relevant sources across more than one type, with important gaps.
- Strong: good coverage across official, editorial/comparison, competitor, and user/community/review sources.
- Robust: 8-10+ relevant sources, multiple types, recent evidence, convergence, and contradictions checked.

Use the right proof standard:
- Official/owned sources: specs, prices, plans, dimensions, weight, battery, compatibility, eligibility, and availability.
- Retail/product feeds: price, availability, specs, and ratings at scale.
- Editorial/comparison: best-fit recommendations, comparisons, and tradeoffs.
- Lab/technical testing: objective performance superlatives and physical/technical superiority.
- Community/user sources: comfort, fit, reliability, shipping reality, service, and real-world use.

Do not treat independent lab testing as universally superior. For price, promo, offer, savings, or value claims, use Commerce editorial / comparison, Independent price audit / comparison table, Retail / product feed, Official pricing page, and Community / forums. Do not use Lab / technical testing for those claims.

Strict source separation:
- Current Web Reality means sources actually searched, retrieved, or returned by web search in this run. Only say what a source currently says if it was actually checked. If no source was checked, say no current web sources were checked.
- Validation Target Map means sources or source types an AI recommendation agent would likely use to validate this claim type. These are targets, not proof that the source currently supports the claim.
- Hypothetical proof language means what validation targets would need to publish. Never present hypothetical proof language as current source reality.
- Counterfactual Source Echo means a simulation of what would happen if validation targets repeated either the claim exactly as written or the recommended upgraded version.

Claim payload, proof, and source echo separation:
- Claim Payload is the structured argument the brand wants the market to understand.
- Proof Requirements are the facts, metrics, comparison set, methodology, caveats, currentness, and supporting artifacts needed to make that payload credible.
- Source Echo is who says the payload, where it appears, and how much authority that source context adds.

Do not force a third-party attribution phrase such as "according to [source]" into the source-ready argument. If an editorial, comparison, review, or community source is the speaker, its authorship provides the attribution. Only use explicit third-party attribution in the separate owned-site citation version when appropriate.

Do not downgrade a claim merely because the payload lacks source attribution when the simulation condition says a third party is echoing it. Do not require a single source to rank a product number one. Convergent evidence across editorial, comparison, review-aggregator, community, competitor, and official sources can make a claim recommendation-useful.

Use source status badges exactly where appropriate: Verified source found, Source target only, Hypothetical proof language, User-supplied claim, Not checked.

Calibration example from controlled tests:
For the query "Which meal kit delivery is the best affordable option with easy-to-make recipes?", the baseline evidence did not put HelloFresh in the consideration set. A light promo claim, "40% off for the first year", still did not put HelloFresh in the consideration set. A structured comparative promo claim with duration, plan assumption, effective price, competitor parity, competitor undercut, shipping caveat, and non-price advantage did put HelloFresh into the consideration set in ChatGPT and Gemini tests, and sometimes made it preferred.

Use this as a general calibration principle only: repetition alone is insufficient. A claim must be understood, trusted, structured, comparative, caveated, and decision-relevant before it can become recommendation-changing.

Category-agnostic promo calibration:
- If a price, savings, discount, or offer claim only states the promotion in broad terms, classify the Claim Structure Status as Mention-worthy unless the current input includes enough decision structure to compare it. Likely recommendation impact if echoed as-is should be low. Explain that it may be mentioned as a promotion but the agent will likely continue recommending the current category leader or budget default because base price, effective price, assumptions, fees, eligibility, duration, competitor comparison, and decision-relevant tradeoffs are missing.
- If a price, savings, discount, or offer claim includes duration, plan/package assumption, effective price or resulting value, competitor parity or undercut, fee/shipping/tax caveat, regular-vs-promotional distinction, eligibility, and a relevant non-price advantage, classify the Claim Structure Status as Consideration-worthy or stronger depending on corroboration. Expected behavior should be that the brand can enter the consideration set and may become a conditional recommendation for the specified scenario, while existing category leaders may remain defaults outside those assumptions.

Do not reuse examples, brands, prices, competitors, or claims from prior calibration examples unless they are explicitly present in the user's current input. Infer the claim type and required proof structure from the current user input only.

If the user enters a legal services claim, healthcare claim, financial claim, tax claim, or other regulated-category claim, do not force it into promo/price logic. Identify the appropriate claim type and evidence standard for that category.

For web-grounded mode, do not treat web search as merely checking whether the claim already exists online. For new claims, identify where AI agents would likely look for validation, what those source targets currently say when discoverable, and what they would need to say for the claim to become consideration-worthy or recommendation-changing. Name discovered source targets when possible, such as category-relevant publications, comparison sites, official pages, data feeds, standards bodies, forums, Reddit communities, creators, or review sources.

Math and measurement guardrail:
When the user provides dimensions, prices, discounts, percentages, or other numbers, preserve the user's exact numbers and recalculate implied percentages only for the metric actually defined. Do not convert a dimension change into a broader "smaller" or "larger" claim unless the claim defines that broader metric, such as diameter, volume, weight, total profile, or surface area.

For size, thickness, weight, volume, comfort, durability, speed, or performance claims:
1. Preserve the stated number.
2. Identify the likely measurement basis.
3. Ask what the number refers to if unclear.
4. Only flag the math as a contradiction if the claim explicitly ties a percentage to a specific calculation that does not match.
5. If multiple interpretations are plausible, say so instead of treating the claim as wrong.

Do not downgrade a claim merely because it uses cleaner consumer-facing language if the brand can support the measurement basis. Downgrade when the metric is undefined, the same metric is calculated inconsistently, the comparison set is unclear, the claim implies a user benefit without evidence, or the claim is unlikely to matter to the target recommendation context.

For size claims, distinguish thickness reduction, diameter reduction, volume reduction, weight reduction, perceived comfort, and user outcome. A technical claim may be mention-worthy for general shoppers but consideration-worthy for a specific user need, such as comfort, sleep wear, low-profile design, aesthetics, or unobtrusive all-day use.

Intent and job-to-be-done calibration:
Infer the specific user need implied by the current claim when possible. Score broad category impact separately from impact for that user need. A claim can be only mention-worthy for the general category but consideration-worthy or a conditional recommendation for a narrow intent it directly supports.

Do not require unrelated criteria to block intent-specific consideration. If a claim directly proves frequency, capacity, cadence, convenience, comfort, or another stated need, treat price, satisfaction, style fit, service, and other criteria as tradeoffs unless the intended recommendation question is about those criteria.

For apparel rental and analogous capacity/cadence claims, distinguish:
- catalog breadth: available styles, brands, designs, or categories
- rotation capacity: how many items a user can receive in a period
- refresh cadence: how often a user can swap or receive items

Use "items" rather than "outfits" unless a checked source explicitly calls them outfits. An item may be a dress, jacket, top, accessory, or another component. Safer language may explain that a number of items can support more outfit rotation without equating items and outfits.

Third-Party Story Map:
Treat source storytelling as a primary output. In web-grounded mode, include up to 10 relevant sources if that many were actually scanned or identified. Separate found-and-scanned sources from weak support, targets, outreach opportunities, community checks, and competitor fact checks. Never claim a source was scanned unless it was actually retrieved in this run.

Claim Routes:
For every claim, generate 2-4 strategically distinct routes that could make it more recommendation-useful. Routes may include category leadership, best fit for a specific user, measurable attribute leadership, best balance, upgrade proof, value, or another route inferred from the current claim. Do not force every route into independent lab testing; choose proof appropriate to the route.

Choose one Recommended Route based on likely recommendation impact, achievability, decision relevance, and available proof. The Argument to Seed and Proof Packet should correspond to that recommended route.

Natural source discovery:
"Who Needs to Say It" must be grounded in natural search behavior that approximates how an AI recommendation agent gathers evidence.
1. Infer a likely consumer recommendation query from the claim, category, and user intent.
2. In web-grounded mode, search naturally around that query, the category, brand, competitors, and decision criteria.
3. Identify sources that actually appear in the evidence layer.
4. Classify those sources by role.
5. State what checked sources currently say.
6. State what they would need to say to strengthen the route.

Do not preload or randomly guess famous publications. A named source not found or scanned in this run must be labeled "Source target only — not found/scanned in this run." Discovered and scanned sources must come before optional targets. Missing source types should describe evidence gaps, not invented publications.

Source Echo Impact Ladder:
For each claim route, simulate Brand site only, Brand site + structured data, One editorial/comparison source, Multiple editorial/comparison sources, Editorial + review aggregator data, Editorial + community corroboration, and Editorial + community contradiction. Use: Ignored, Brand context only, Mention-worthy, Consideration-worthy, Strong consideration, Conditional recommendation, Recommendation-changing.
""".strip()


ECHO_STAGE_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "stage": {
            "type": "string",
            "enum": [
                "Claim as written",
                "If echoed by owned site only",
                "If echoed by one commerce/editorial source",
                "If echoed by multiple commerce/editorial sources",
                "If echoed by sources plus Reddit/community corroboration",
                "If echoed by sources but contradicted by Reddit/community",
            ],
        },
        "impact_label": {
            "type": "string",
            "enum": IMPACT_LEVELS,
        },
        "source_status_badge": {
            "type": "string",
            "enum": [
                "Verified source found",
                "Source target only",
                "Hypothetical proof language",
                "User-supplied claim",
                "Not checked",
            ],
        },
        "likely_agent_behavior": {"type": "string"},
        "why": {"type": "string"},
        "structural_improvement_needed": {"type": "string"},
        "source_validation_needed": {"type": "string"},
    },
    "required": [
        "stage",
        "impact_label",
        "source_status_badge",
        "likely_agent_behavior",
        "why",
        "structural_improvement_needed",
        "source_validation_needed",
    ],
}


DISCOVERED_SOURCE_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "source_name": {"type": "string"},
        "url": {"type": "string"},
        "discovery_path_or_query": {"type": "string"},
        "source_type": {
            "type": "string",
            "enum": [
                "Editorial / comparison",
                "Commerce editorial / comparison",
                "Independent price audit / comparison table",
                "Lab / technical testing",
                "User review aggregator",
                "Community / forums",
                "Retail / product feed",
                "Official pricing page",
                "Owned / official specs",
                "Competitor fact-check",
                "Other",
            ],
        },
        "source_status": {
            "type": "string",
            "enum": [
                "Scanned",
                "Discovered, not scanned",
                "Source target only — not found/scanned in this run",
                "Missing source type",
                "Not checked",
            ],
        },
        "why_relevant_to_recommendation_ecosystem": {"type": "string"},
        "what_it_currently_says": {"type": "string"},
        "current_stance": {
            "type": "string",
            "enum": [
                "Supports",
                "Partially supports",
                "Contradicts",
                "Does not mention",
                "Not scanned",
                "Not checked",
            ],
        },
        "what_it_would_need_to_say": {"type": "string"},
        "likely_impact_if_echoed_upgraded_claim": {
            "type": "string",
            "enum": IMPACT_LEVELS,
        },
        "source_role": {"type": "string"},
        "can_brand_influence": {
            "type": "string",
            "enum": ["Yes", "Partially", "Limited", "No"],
        },
        "contact_action_priority": {
            "type": "string",
            "enum": ["High", "Medium", "Low"],
        },
        "recommended_brand_action": {"type": "string"},
    },
    "required": [
        "source_name",
        "url",
        "discovery_path_or_query",
        "source_type",
        "source_status",
        "why_relevant_to_recommendation_ecosystem",
        "what_it_currently_says",
        "current_stance",
        "what_it_would_need_to_say",
        "likely_impact_if_echoed_upgraded_claim",
        "source_role",
        "can_brand_influence",
        "contact_action_priority",
        "recommended_brand_action",
    ],
}


ANALYSIS_SCHEMA = {
    "name": "claim_examiner_analysis",
    "schema": {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "claim_routes": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "route_name": {"type": "string"},
                        "what_current_claim_is_trying_to_become": {"type": "string"},
                        "stronger_source_ready_claim": {"type": "string"},
                        "user_intent_or_recommendation_lane": {"type": "string"},
                        "key_decision_criteria": {"type": "array", "items": {"type": "string"}},
                        "proof_needed": {"type": "array", "items": {"type": "string"}},
                        "best_source_types": {"type": "array", "items": {"type": "string"}},
                        "likely_impact_if_validated": {
                            "type": "string",
                            "enum": IMPACT_LEVELS,
                        },
                        "main_tradeoffs_or_risks": {"type": "array", "items": {"type": "string"}},
                    },
                    "required": [
                        "route_name",
                        "what_current_claim_is_trying_to_become",
                        "stronger_source_ready_claim",
                        "user_intent_or_recommendation_lane",
                        "key_decision_criteria",
                        "proof_needed",
                        "best_source_types",
                        "likely_impact_if_validated",
                        "main_tradeoffs_or_risks",
                    ],
                },
            },
            "recommended_route": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "route_name": {"type": "string"},
                    "why_this_route": {"type": "string"},
                    "recommendation_potential": {"type": "string"},
                },
                "required": ["route_name", "why_this_route", "recommendation_potential"],
            },
            "source_discovery_summary": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "search_intent_used": {"type": "string"},
                    "queries_run": {"type": "array", "items": {"type": "string"}},
                    "sources_discovered": {"type": "integer"},
                    "sources_scanned": {"type": "integer"},
                    "source_mix": {"type": "array", "items": {"type": "string"}},
                    "gaps_in_coverage": {"type": "array", "items": {"type": "string"}},
                    "key_gap": {"type": "string"},
                    "highest_leverage_source_opportunity": {"type": "string"},
                    "evidence_layer_summary": {"type": "string"},
                },
                "required": [
                    "search_intent_used",
                    "queries_run",
                    "sources_discovered",
                    "sources_scanned",
                    "source_mix",
                    "gaps_in_coverage",
                    "key_gap",
                    "highest_leverage_source_opportunity",
                    "evidence_layer_summary",
                ],
            },
            "who_needs_to_say_it": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "discovered_sources": {"type": "array", "items": DISCOVERED_SOURCE_SCHEMA},
                    "scanned_sources": {"type": "array", "items": DISCOVERED_SOURCE_SCHEMA},
                    "source_targets_only": {"type": "array", "items": DISCOVERED_SOURCE_SCHEMA},
                    "missing_source_types": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "additionalProperties": False,
                            "properties": {
                                "source_type": {"type": "string"},
                                "why_missing_evidence_matters": {"type": "string"},
                                "recommended_search_or_action": {"type": "string"},
                            },
                            "required": [
                                "source_type",
                                "why_missing_evidence_matters",
                                "recommended_search_or_action",
                            ],
                        },
                    },
                },
                "required": [
                    "discovered_sources",
                    "scanned_sources",
                    "source_targets_only",
                    "missing_source_types",
                ],
            },
            "contact_priority_list": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "priority_rank": {"type": "integer"},
                        "source_or_source_type": {"type": "string"},
                        "why_it_matters": {"type": "string"},
                        "what_to_give_or_do": {"type": "string"},
                        "expected_impact_if_successful": {
                            "type": "string",
                            "enum": IMPACT_LEVELS,
                        },
                    },
                    "required": [
                        "priority_rank",
                        "source_or_source_type",
                        "why_it_matters",
                        "what_to_give_or_do",
                        "expected_impact_if_successful",
                    ],
                },
            },
            "source_echo_impact_ladder": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "route_name": {"type": "string"},
                        "stages": {
                            "type": "array",
                            "items": {
                                "type": "object",
                                "additionalProperties": False,
                                "properties": {
                                    "echo_stage": {
                                        "type": "string",
                                        "enum": [
                                            "Brand site only",
                                            "Brand site + structured data",
                                            "One editorial/comparison source",
                                            "Multiple editorial/comparison sources",
                                            "Editorial + review aggregator data",
                                            "Editorial + community corroboration",
                                            "Editorial + community contradiction",
                                        ],
                                    },
                                    "impact": {
                                        "type": "string",
                                        "enum": IMPACT_LEVELS,
                                    },
                                    "why": {"type": "string"},
                                },
                                "required": ["echo_stage", "impact", "why"],
                            },
                        },
                    },
                    "required": ["route_name", "stages"],
                },
            },
            "intent_match": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "inferred_user_intent": {"type": "string"},
                    "job_to_be_done": {"type": "string"},
                    "how_claim_maps_to_intent": {"type": "string"},
                    "intent_confidence_note": {"type": "string"},
                },
                "required": [
                    "inferred_user_intent",
                    "job_to_be_done",
                    "how_claim_maps_to_intent",
                    "intent_confidence_note",
                ],
            },
            "general_category_impact": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "status": {
                        "type": "string",
                        "enum": IMPACT_LEVELS,
                    },
                    "reason": {"type": "string"},
                },
                "required": ["status", "reason"],
            },
            "intent_specific_impact": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "status": {
                        "type": "string",
                        "enum": IMPACT_LEVELS,
                    },
                    "intent": {"type": "string"},
                    "reason": {"type": "string"},
                    "tradeoff_note": {"type": "string"},
                },
                "required": ["status", "intent", "reason", "tradeoff_note"],
            },
            "decision_criteria_detected": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "primary_decision_criterion": {"type": "string"},
                    "secondary_decision_criteria": {"type": "array", "items": {"type": "string"}},
                    "tradeoffs": {"type": "array", "items": {"type": "string"}},
                    "missing_proof": {"type": "array", "items": {"type": "string"}},
                    "criterion_support": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "additionalProperties": False,
                            "properties": {
                                "criterion": {"type": "string"},
                                "support_type": {
                                    "type": "string",
                                    "enum": ["direct", "indirect", "not supported"],
                                },
                                "evidence_from_claim": {"type": "string"},
                            },
                            "required": ["criterion", "support_type", "evidence_from_claim"],
                        },
                    },
                    "variety_dimensions": {
                        "type": "object",
                        "additionalProperties": False,
                        "properties": {
                            "catalog_breadth": {"type": "string"},
                            "rotation_capacity": {"type": "string"},
                            "refresh_cadence": {"type": "string"},
                        },
                        "required": ["catalog_breadth", "rotation_capacity", "refresh_cadence"],
                    },
                },
                "required": [
                    "primary_decision_criterion",
                    "secondary_decision_criteria",
                    "tradeoffs",
                    "missing_proof",
                    "criterion_support",
                    "variety_dimensions",
                ],
            },
            "simulated_recommendation": {"type": "string"},
            "argument_to_seed": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "source_ready_argument": {"type": "string"},
                    "owned_site_citation_version": {"type": "string"},
                    "claim_payload_note": {"type": "string"},
                },
                "required": [
                    "source_ready_argument",
                    "owned_site_citation_version",
                    "claim_payload_note",
                ],
            },
            "proof_packet": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "metrics": {"type": "array", "items": {"type": "string"}},
                    "comparison_set": {"type": "array", "items": {"type": "string"}},
                    "sourceable_facts": {"type": "array", "items": {"type": "string"}},
                    "caveats": {"type": "array", "items": {"type": "string"}},
                    "methodology": {"type": "array", "items": {"type": "string"}},
                    "currentness": {"type": "array", "items": {"type": "string"}},
                    "artifacts_needed": {"type": "array", "items": {"type": "string"}},
                },
                "required": [
                    "metrics",
                    "comparison_set",
                    "sourceable_facts",
                    "caveats",
                    "methodology",
                    "currentness",
                    "artifacts_needed",
                ],
            },
            "source_echo_simulation": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "source_type": {
                            "type": "string",
                            "enum": [
                                "owned site",
                                "third-party editorial/review",
                                "commerce comparison",
                                "user review aggregator",
                                "community/forum",
                                "competitor source",
                                "official/regulatory/standards source",
                            ],
                        },
                        "how_argument_would_read": {"type": "string"},
                        "authority_upgrade": {"type": "string"},
                        "surrounding_evidence_required": {"type": "string"},
                        "likely_recommendation_impact": {
                            "type": "string",
                            "enum": IMPACT_LEVELS,
                        },
                    },
                    "required": [
                        "source_type",
                        "how_argument_would_read",
                        "authority_upgrade",
                        "surrounding_evidence_required",
                        "likely_recommendation_impact",
                    ],
                },
            },
            "confidence_source_coverage": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "status": {"type": "string", "enum": EVIDENCE_COVERAGE_LEVELS},
                    "note": {"type": "string"},
                },
                "required": ["status", "note"],
            },
            "claim_structure_status": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "status": {
                        "type": "string",
                        "enum": IMPACT_LEVELS,
                    },
                    "likely_recommendation_impact_if_echoed_as_is": {"type": "string"},
                    "why": {"type": "string"},
                    "expected_agent_behavior": {"type": "string"},
                },
                "required": [
                    "status",
                    "likely_recommendation_impact_if_echoed_as_is",
                    "why",
                    "expected_agent_behavior",
                ],
            },
            "current_web_reality": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "source_name": {"type": "string"},
                        "url": {"type": "string"},
                        "source_type": {"type": "string"},
                        "status": {
                            "type": "string",
                            "enum": ["found", "not found", "inaccessible", "not checked"],
                        },
                        "source_status_badge": {
                            "type": "string",
                            "enum": [
                                "Verified source found",
                                "Source target only",
                                "Hypothetical proof language",
                                "User-supplied claim",
                                "Not checked",
                            ],
                        },
                        "what_it_currently_says": {"type": "string"},
                        "supports_claim": {
                            "type": "string",
                            "enum": ["yes", "partially", "no", "does not mention"],
                        },
                        "notes_or_caveats": {"type": "string"},
                    },
                    "required": [
                        "source_name",
                        "url",
                        "source_type",
                        "status",
                        "source_status_badge",
                        "what_it_currently_says",
                        "supports_claim",
                        "notes_or_caveats",
                    ],
                },
            },
            "recommendation_impact": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "status": {
                        "type": "string",
                        "enum": IMPACT_LEVELS,
                    },
                    "why": {"type": "string"},
                    "what_would_change_the_recommendation": {"type": "string"},
                },
                "required": ["status", "why", "what_would_change_the_recommendation"],
            },
            "claim_elements": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "element": {"type": "string"},
                        "classification": {
                            "type": "string",
                            "enum": [
                                "usable evidence",
                                "potentially useful but missing context",
                                "brand context only",
                                "fluff / unsupported conclusion",
                                "risky or ambiguous",
                            ],
                        },
                        "what_is_missing": {"type": "string"},
                        "why_it_matters": {"type": "string"},
                    },
                    "required": [
                        "element",
                        "classification",
                        "what_is_missing",
                        "why_it_matters",
                    ],
                },
            },
            "math_logic_audit": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "numbers_detected": {"type": "string"},
                    "percentages": {"type": "array", "items": {"type": "string"}},
                    "savings_or_price_claims": {"type": "array", "items": {"type": "string"}},
                    "times_claims": {"type": "array", "items": {"type": "string"}},
                    "size_change_claims": {"type": "array", "items": {"type": "string"}},
                    "math_questions_to_resolve": {"type": "array", "items": {"type": "string"}},
                    "wording_accuracy_risk": {"type": "string"},
                },
                "required": [
                    "numbers_detected",
                    "percentages",
                    "savings_or_price_claims",
                    "times_claims",
                    "size_change_claims",
                    "math_questions_to_resolve",
                    "wording_accuracy_risk",
                ],
            },
            "stronger_agent_ready_argument": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "structured_rewrite": {"type": "string"},
                    "path_from_marketing_to_decision_proof": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "additionalProperties": False,
                            "properties": {
                                "move": {"type": "string"},
                                "why_it_makes_the_claim_more_agent_usable": {"type": "string"},
                            },
                            "required": [
                                "move",
                                "why_it_makes_the_claim_more_agent_usable",
                            ],
                        },
                    },
                    "decision_proof_template": {"type": "string"},
                },
                "required": [
                    "structured_rewrite",
                    "path_from_marketing_to_decision_proof",
                    "decision_proof_template",
                ],
            },
            "recommended_upgrade": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "upgraded_claim": {"type": "string"},
                    "required_structure": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "additionalProperties": False,
                            "properties": {
                                "component": {"type": "string"},
                                "why_it_is_needed": {"type": "string"},
                                "example_placeholder": {"type": "string"},
                            },
                            "required": [
                                "component",
                                "why_it_is_needed",
                                "example_placeholder",
                            ],
                        },
                    },
                    "sourceable_proof_structure": {"type": "string"},
                },
                "required": [
                    "upgraded_claim",
                    "required_structure",
                    "sourceable_proof_structure",
                ],
            },
            "counterfactual_source_echo": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "exact_claim_as_written": {
                        "type": "array",
                        "items": ECHO_STAGE_SCHEMA,
                    },
                    "recommended_upgraded_version": {
                        "type": "array",
                        "items": ECHO_STAGE_SCHEMA,
                    },
                },
                "required": ["exact_claim_as_written", "recommended_upgraded_version"],
            },
            "validation_target_map": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "source_name_or_type": {"type": "string"},
                        "source_status_badge": {
                            "type": "string",
                            "enum": [
                                "Verified source found",
                                "Source target only",
                                "Hypothetical proof language",
                                "User-supplied claim",
                                "Not checked",
                            ],
                        },
                        "why_an_ai_agent_would_care": {"type": "string"},
                        "evidence_that_would_need_to_appear": {"type": "string"},
                        "can_brand_ethically_influence_it": {"type": "string"},
                        "recommended_action": {"type": "string"},
                    },
                    "required": [
                        "source_name_or_type",
                        "source_status_badge",
                        "why_an_ai_agent_would_care",
                        "evidence_that_would_need_to_appear",
                        "can_brand_ethically_influence_it",
                        "recommended_action",
                    ],
                },
            },
            "third_party_story_map": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "coverage_note": {"type": "string"},
                    "sources": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "additionalProperties": False,
                            "properties": {
                                "source_name": {"type": "string"},
                                "url": {"type": "string"},
                                "source_type": {"type": "string"},
                                "source_group": {
                                    "type": "string",
                                    "enum": [
                                        "Editorial / comparison",
                                        "User review aggregators",
                                        "Community / forums",
                                        "Owned site",
                                        "Competitor fact-check",
                                        "Other",
                                    ],
                                },
                                "scan_status": {
                                    "type": "string",
                                    "enum": [
                                        "found and scanned",
                                        "found but weak",
                                        "source target only",
                                        "not scanned",
                                    ],
                                },
                                "current_stance": {
                                    "type": "string",
                                    "enum": [
                                        "supports",
                                        "partially supports",
                                        "contradicts",
                                        "does not mention",
                                        "not scanned",
                                    ],
                                },
                                "what_it_currently_says": {"type": "string"},
                                "what_it_would_need_to_say": {"type": "string"},
                                "why_an_ai_agent_would_care": {"type": "string"},
                                "recommended_brand_action": {"type": "string"},
                            },
                            "required": [
                                "source_name",
                                "url",
                                "source_type",
                                "source_group",
                                "scan_status",
                                "current_stance",
                                "what_it_currently_says",
                                "what_it_would_need_to_say",
                                "why_an_ai_agent_would_care",
                                "recommended_brand_action",
                            ],
                        },
                    },
                },
                "required": ["coverage_note", "sources"],
            },
            "source_validation_plan": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "source_role": {
                            "type": "string",
                            "enum": [
                                "owned site / brand verification",
                                "official authority / regulator / standards body",
                                "third-party expert review",
                                "commerce editorial / affiliate comparison",
                                "product data feed / pricing source",
                                "creator / YouTube demonstration",
                                "Reddit/community reality check",
                            ],
                        },
                        "why_an_ai_agent_would_care": {"type": "string"},
                        "target_examples_or_discovered_sources": {
                            "type": "array",
                            "items": {"type": "string"},
                        },
                        "what_sources_need_to_say": {"type": "string"},
                        "exact_evidence_that_should_appear": {"type": "string"},
                        "can_brand_ethically_influence_it": {"type": "string"},
                        "suggested_action": {"type": "string"},
                    },
                    "required": [
                        "source_role",
                        "why_an_ai_agent_would_care",
                        "target_examples_or_discovered_sources",
                        "what_sources_need_to_say",
                        "exact_evidence_that_should_appear",
                        "can_brand_ethically_influence_it",
                        "suggested_action",
                    ],
                },
            },
            "claim_verdict": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "safe_to_repeat_as_fact": {"type": "string"},
                    "safe_to_repeat_only_as_brand_claim": {"type": "string"},
                    "needs_substantiation": {"type": "string"},
                    "likely_ignored_by_ai_agent": {"type": "string"},
                    "could_become_recommendation_changing_with_proof": {"type": "string"},
                    "summary": {"type": "string"},
                },
                "required": [
                    "safe_to_repeat_as_fact",
                    "safe_to_repeat_only_as_brand_claim",
                    "needs_substantiation",
                    "likely_ignored_by_ai_agent",
                    "could_become_recommendation_changing_with_proof",
                    "summary",
                ],
            },
            "claim_type": {
                "type": "array",
                "items": {"type": "string"},
            },
            "what_ai_would_say_today": {"type": "string"},
            "missing_evidence": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "piece": {"type": "string"},
                        "why_it_matters": {"type": "string"},
                        "example_needed": {"type": "string"},
                    },
                    "required": ["piece", "why_it_matters", "example_needed"],
                },
            },
            "triangulation_plan": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "surface": {"type": "string"},
                        "what_to_publish_or_collect": {"type": "string"},
                        "why_an_ai_agent_would_care": {"type": "string"},
                    },
                    "required": [
                        "surface",
                        "what_to_publish_or_collect",
                        "why_an_ai_agent_would_care",
                    ],
                },
            },
            "safer_claim_rewrites": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "safe_marketing_language": {"type": "string"},
                    "more_substantiated_version": {"type": "string"},
                    "recommendation_ready_version_assuming_evidence_exists": {"type": "string"},
                },
                "required": [
                    "safe_marketing_language",
                    "more_substantiated_version",
                    "recommendation_ready_version_assuming_evidence_exists",
                ],
            },
            "agent_readable_evidence_block": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "claim": {"type": "string"},
                    "claim_type": {"type": "array", "items": {"type": "string"}},
                    "audience": {"type": "string"},
                    "comparison_set": {"type": "string"},
                    "metric": {"type": "string"},
                    "test_method": {"type": "string"},
                    "evidence_needed": {"type": "array", "items": {"type": "string"}},
                    "caveats": {"type": "array", "items": {"type": "string"}},
                    "status": {"type": "string"},
                },
                "required": [
                    "claim",
                    "claim_type",
                    "audience",
                    "comparison_set",
                    "metric",
                    "test_method",
                    "evidence_needed",
                    "caveats",
                    "status",
                ],
            },
            "red_team": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "risk": {"type": "string"},
                        "how_it_could_be_attacked_or_ignored": {"type": "string"},
                        "how_to_neutralize": {"type": "string"},
                    },
                    "required": [
                        "risk",
                        "how_it_could_be_attacked_or_ignored",
                        "how_to_neutralize",
                    ],
                },
            },
            "recommendation_readiness": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "label": {"type": "string"},
                    "rationale": {"type": "string"},
                    "next_best_action": {"type": "string"},
                },
                "required": ["label", "rationale", "next_best_action"],
            },
            "citations": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "title": {"type": "string"},
                        "url": {"type": "string"},
                        "note": {"type": "string"},
                    },
                    "required": ["title", "url", "note"],
                },
            },
        },
        "required": [
            "claim_routes",
            "recommended_route",
            "source_discovery_summary",
            "who_needs_to_say_it",
            "source_echo_impact_ladder",
            "contact_priority_list",
            "intent_match",
            "general_category_impact",
            "intent_specific_impact",
            "decision_criteria_detected",
            "simulated_recommendation",
            "argument_to_seed",
            "proof_packet",
            "source_echo_simulation",
            "confidence_source_coverage",
            "claim_structure_status",
            "current_web_reality",
            "recommendation_impact",
            "claim_elements",
            "math_logic_audit",
            "stronger_agent_ready_argument",
            "recommended_upgrade",
            "counterfactual_source_echo",
            "validation_target_map",
            "third_party_story_map",
            "source_validation_plan",
            "claim_verdict",
            "claim_type",
            "what_ai_would_say_today",
            "missing_evidence",
            "triangulation_plan",
            "safer_claim_rewrites",
            "agent_readable_evidence_block",
            "red_team",
            "recommendation_readiness",
            "citations",
        ],
    },
    "strict": True,
}

QUERY_PLAN_SCHEMA = {
    "name": "source_discovery_query_plan",
    "schema": {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "search_intent": {"type": "string"},
            "queries": {
                "type": "array",
                "items": {"type": "string"},
            },
        },
        "required": ["search_intent", "queries"],
    },
    "strict": True,
}


def load_api_key() -> tuple[str, str]:
    load_dotenv()
    environment_key = os.getenv("OPENAI_API_KEY", "").strip()
    if environment_key:
        return environment_key, "environment"

    try:
        secrets_key = str(st.secrets.get("OPENAI_API_KEY", "")).strip()
    except Exception:
        secrets_key = ""
    if secrets_key:
        return secrets_key, "streamlit secrets"

    return "", ""


def get_text_from_upload(uploaded_file) -> str:
    suffix = uploaded_file.name.lower().split(".")[-1]
    data = uploaded_file.read()

    if suffix in {"txt", "md", "csv"}:
        return data.decode("utf-8", errors="replace")

    if suffix == "pdf":
        if PdfReader is None:
            return "[PDF uploaded, but pypdf is not installed.]"
        reader = PdfReader(BytesIO(data))
        pages = [page.extract_text() or "" for page in reader.pages]
        return "\n\n".join(pages).strip()

    if suffix == "docx":
        if Document is None:
            return "[DOCX uploaded, but python-docx is not installed.]"
        document = Document(BytesIO(data))
        return "\n".join(paragraph.text for paragraph in document.paragraphs).strip()

    return f"[Unsupported file type: {uploaded_file.name}]"


def compact_source_bundle(source_discovery: dict[str, Any] | None) -> dict[str, Any]:
    if not source_discovery:
        return {
            "search_intent": "No web discovery performed.",
            "queries": [],
            "sources": [],
        }
    compact_sources = []
    for source in source_discovery.get("sources", [])[:MAX_DISCOVERY_SOURCES]:
        compact_sources.append(
            {
                "title": source.get("title", "Web source"),
                "url": source.get("url", ""),
                "queries": source.get("discovery_queries", [])[:2],
                "citation_backed": bool(source.get("citation_backed")),
                "summary": " ".join(source.get("summaries", []))[:MAX_DISCOVERY_SUMMARY_CHARS],
            }
        )
    return {
        "search_intent": source_discovery.get("search_intent", ""),
        "queries": source_discovery.get("queries", [])[:4],
        "sources": compact_sources,
    }


def build_user_prompt(
    inputs: dict[str, Any],
    evidence_text: str,
    source_discovery: dict[str, Any] | None = None,
) -> str:
    source_bundle = compact_source_bundle(source_discovery)
    return f"""
Analyze the claim using the system instructions and return the required structured result.

Date: {date.today().isoformat()}
Research mode: {inputs["research_mode"]}
Brand/product: {inputs["brand"] or "[Not provided]"}
Category: {inputs["category"]}
Claim: {inputs["claim"]}
Known facts: {inputs["known_facts"] or "[Not provided]"}
Competitors: {inputs["competitors"] or "[Not provided]"}
Uploaded evidence: {(evidence_text or "[No uploaded evidence]")[:6000]}
Verified web discovery: {json.dumps(source_bundle, separators=(",", ":"))}

Grounding constraints:
- Only sources in Verified web discovery may be named as discovered, scanned, current, or cited.
- Never invent source names or source-attributed currentness.
- Unverified targets must be generic source types.
- Treat unsupported claim facts as user-supplied.
- Generate 2-4 distinct claim routes and choose the strongest route.
- Keep the result concise: use 2-3 claim routes, no more than 6 source entries, no more than 5 missing-evidence items, and short paragraphs.
""".strip()


def extract_json_from_response(response) -> dict[str, Any]:
    text = getattr(response, "output_text", "") or ""
    if not text:
        for item in getattr(response, "output", []) or []:
            for content in getattr(item, "content", []) or []:
                if getattr(content, "type", None) == "output_text":
                    text += getattr(content, "text", "")
    return json.loads(text)


def extract_citations_from_response(response) -> list[dict[str, str]]:
    citations = []
    seen = set()
    for item in getattr(response, "output", []) or []:
        for content in object_value(item, "content", []) or []:
            for annotation in object_value(content, "annotations", []) or []:
                url = object_value(annotation, "url")
                if url and url not in seen:
                    seen.add(url)
                    citations.append(
                        {
                            "title": object_value(annotation, "title", "") or "Source",
                            "url": url,
                            "note": "Source used during web-grounded analysis.",
                        }
                    )
    return citations


def extract_web_queries_from_response(response) -> list[str]:
    queries = []
    seen = set()
    for item in getattr(response, "output", []) or []:
        if getattr(item, "type", None) != "web_search_call":
            continue
        action = getattr(item, "action", None)
        candidates = [
            getattr(action, "query", None),
            *(getattr(action, "queries", None) or []),
        ]
        for query in candidates:
            if query and query not in seen:
                seen.add(query)
                queries.append(query)
    return queries


def object_value(obj: Any, key: str, default=None):
    if isinstance(obj, dict):
        return obj.get(key, default)
    return getattr(obj, key, default)


def extract_web_activity(response) -> dict[str, Any]:
    queries = []
    sources = []
    seen_queries = set()
    seen_sources = set()
    call_count = 0
    opened_page_count = 0

    for item in getattr(response, "output", []) or []:
        if object_value(item, "type") != "web_search_call":
            continue
        call_count += 1
        action = object_value(item, "action")
        action_type = object_value(action, "type", "")
        if action_type == "open_page":
            opened_page_count += 1

        candidates = [
            object_value(action, "query"),
            *(object_value(action, "queries", []) or []),
        ]
        for query in candidates:
            if query and query not in seen_queries:
                seen_queries.add(query)
                queries.append(query)

        for source in object_value(action, "sources", []) or []:
            url = object_value(source, "url", "")
            title = object_value(source, "title", "") or object_value(source, "name", "")
            if url and url not in seen_sources:
                seen_sources.add(url)
                sources.append({"title": title or "Web source", "url": url})

    return {
        "web_search_call_count": call_count,
        "opened_page_count": opened_page_count,
        "queries": queries,
        "sources": sources,
    }


def fallback_discovery_queries(inputs: dict[str, Any]) -> dict[str, Any]:
    brand = inputs["brand"].strip()
    category = inputs["category"].strip()
    competitors = re.sub(r"[,;/]+", " ", inputs["competitors"]).strip()
    claim_words = re.findall(r"[\w$%.+-]+", inputs["claim"])
    claim_phrase = " ".join(claim_words[:18])
    brand_set = " ".join(part for part in [brand, competitors] if part)
    queries = [
        f"best {category} comparison {brand_set}".strip(),
        f"{brand} {claim_phrase}".strip(),
        f"{brand_set} {category} comparison".strip(),
        f"{category} price features comparison {date.today().year}".strip(),
        f"Reddit {brand_set} {category} reviews".strip(),
    ]
    deduped = []
    for query in queries:
        normalized = " ".join(query.split())
        if normalized and normalized not in deduped:
            deduped.append(normalized)
    return {
        "search_intent": (
            f"Find current recommendation evidence about {brand} in {category}, "
            "including relevant competitors and the decision criteria in the claim."
        ),
        "queries": deduped[:4],
    }


def generate_source_query_plan(client: OpenAI, inputs: dict[str, Any]) -> dict[str, Any]:
    prompt = f"""
Create 4 to 6 short, natural web search queries for source discovery.

Brand/product: {inputs["brand"]}
Category: {inputs["category"]}
Claim: {inputs["claim"]}
Competitors: {inputs["competitors"] or "[Not provided]"}
Known facts: {inputs["known_facts"] or "[Not provided]"}

The queries should approximate how a consumer or AI recommendation agent would research:
- the likely recommendation question
- the brand and claim
- competitor comparison
- the main decision criteria
- community reality checks

Each query must be a normal search-engine query, not an instruction or analysis prompt.
Do not include quotation marks around the whole query.
""".strip()
    try:
        response = client.responses.create(
            model=MODEL,
            input=prompt,
            temperature=0.1,
            text={"format": {"type": "json_schema", **QUERY_PLAN_SCHEMA}},
        )
        plan = extract_json_from_response(response)
        queries = [" ".join(query.split()) for query in plan.get("queries", [])]
        queries = [query for query in queries if 3 <= len(query.split()) <= 24]
        queries = list(dict.fromkeys(queries))
        if 4 <= len(queries) <= 8:
            return {
                "search_intent": plan["search_intent"],
                "queries": queries,
            }
    except Exception:
        pass
    return fallback_discovery_queries(inputs)


def run_source_discovery(client: OpenAI, inputs: dict[str, Any]) -> dict[str, Any]:
    plan = fallback_discovery_queries(inputs)
    aggregate_queries = []
    aggregate_sources: dict[str, dict[str, Any]] = {}
    summaries = []
    citations = []
    citation_urls = set()
    web_search_call_count = 0
    opened_page_count = 0
    errors = []

    discovery_prompt = (
        "Research the following short recommendation queries. Run the searches needed, "
        "then give a concise source-backed synthesis. Prioritize official, comparison, "
        "competitor, and community evidence.\n\n"
        + "\n".join(f"- {query}" for query in plan["queries"])
    )
    try:
        response = client.responses.create(
            model=DISCOVERY_MODEL,
            input=discovery_prompt,
            tools=[{"type": "web_search"}],
            tool_choice="required",
            include=["web_search_call.action.sources"],
            max_output_tokens=1800,
        )
    except Exception as exc:
        errors.append(f"{type(exc).__name__}: {exc}")
        response = None

    if response is not None:
        activity = extract_web_activity(response)
        response_citations = extract_citations_from_response(response)
        aggregate_queries = activity["queries"] or plan["queries"]
        web_search_call_count = activity["web_search_call_count"]
        opened_page_count = activity["opened_page_count"]
        summary_text = (getattr(response, "output_text", "") or "").strip()
        if summary_text:
            summaries.append(
                {
                    "query": "Combined discovery",
                    "summary": summary_text[:MAX_DISCOVERY_SUMMARY_CHARS],
                }
            )

        for citation in response_citations:
            url = normalize_url(citation["url"])
            citation_urls.add(url)
            if url not in {normalize_url(item["url"]) for item in citations}:
                citations.append(citation)

        for source in (activity["sources"] + response_citations)[:MAX_DISCOVERY_SOURCES]:
            url = normalize_url(source.get("url", ""))
            if not url:
                continue
            existing = aggregate_sources.setdefault(
                url,
                {
                    "title": source.get("title") or "Web source",
                    "url": source["url"],
                    "discovery_queries": plan["queries"][:2],
                    "summaries": [],
                    "citation_backed": False,
                    "opened_or_scanned": False,
                },
            )
            if summary_text and not existing["summaries"]:
                existing["summaries"].append(summary_text[:MAX_DISCOVERY_SUMMARY_CHARS])
            if url in citation_urls:
                existing["citation_backed"] = True
                existing["opened_or_scanned"] = True

    sources = list(aggregate_sources.values())[:MAX_DISCOVERY_SOURCES]
    return {
        "search_intent": plan["search_intent"],
        "planned_queries": plan["queries"],
        "queries": aggregate_queries,
        "sources": sources,
        "summaries": summaries,
        "citations": citations,
        "web_search_call_count": web_search_call_count,
        "opened_page_count": opened_page_count,
        "errors": errors,
    }


def normalize_url(url: str) -> str:
    return (url or "").strip().rstrip("/")


def canonical_impact(value: str) -> str:
    normalized = (value or "").strip().lower()
    aliases = {
        "ignored": "Ignored",
        "brand context only": "Brand context only",
        "mention-worthy": "Mention-worthy",
        "consideration-worthy": "Consideration-worthy",
        "strong consideration": "Strong consideration",
        "conditional recommendation": "Conditional recommendation",
        "recommendation-changing": "Recommendation-changing",
    }
    return aliases.get(normalized, value if value in IMPACT_LEVELS else "Mention-worthy")


def evidence_coverage(
    scanned_count: int,
    source_mix: list[str],
    web_search_call_count: int,
) -> tuple[str, str]:
    if web_search_call_count == 0:
        return "Not checked", "No web search call completed, so source coverage was not checked."
    if scanned_count <= 2:
        return "Low", f"{scanned_count} source(s) scanned; coverage is very limited."
    if scanned_count <= 4:
        return "Partial", f"{scanned_count} sources scanned with major corroboration gaps."
    if scanned_count <= 7:
        return "Moderate", f"{scanned_count} sources across {len(source_mix)} source type(s); important gaps remain."
    if scanned_count < 10 or len(source_mix) < 4:
        return "Strong", f"{scanned_count} sources provide good cross-source coverage."
    return "Robust", f"{scanned_count} sources across multiple evidence types with broad corroboration."


def normalize_impact_labels(analysis: dict[str, Any]) -> None:
    for key in ["claim_structure_status", "general_category_impact", "intent_specific_impact", "recommendation_impact"]:
        analysis[key]["status"] = canonical_impact(analysis[key]["status"])
    for route in analysis["claim_routes"]:
        route["likely_impact_if_validated"] = canonical_impact(route["likely_impact_if_validated"])
    for route in analysis["source_echo_impact_ladder"]:
        for stage in route["stages"]:
            stage["impact"] = canonical_impact(stage["impact"])
    for source in (
        analysis["who_needs_to_say_it"]["discovered_sources"]
        + analysis["who_needs_to_say_it"]["scanned_sources"]
        + analysis["who_needs_to_say_it"]["source_targets_only"]
    ):
        source["likely_impact_if_echoed_upgraded_claim"] = canonical_impact(
            source["likely_impact_if_echoed_upgraded_claim"]
        )
    for item in analysis["contact_priority_list"]:
        item["expected_impact_if_successful"] = canonical_impact(
            item["expected_impact_if_successful"]
        )
    for echo in analysis["source_echo_simulation"]:
        echo["likely_recommendation_impact"] = canonical_impact(echo["likely_recommendation_impact"])
    for group_name, group in analysis["counterfactual_source_echo"].items():
        for index, stage in enumerate(group):
            stage["impact_label"] = canonical_impact(stage["impact_label"])
            stage["source_status_badge"] = (
                "User-supplied claim"
                if group_name == "exact_claim_as_written" and index == 0
                else "Hypothetical proof language"
            )


def normalize_price_source_types(analysis: dict[str, Any], claim_text: str) -> None:
    claim_signal = (" ".join(analysis.get("claim_type", [])) + " " + claim_text).lower()
    price_claim = any(
        token in claim_signal
        for token in ["price", "offer", "promo", "value", "saving", "discount", "% off", "under $"]
    )
    if not price_claim:
        return
    replacements = {
        "Lab / technical testing": "Independent price audit / comparison table",
        "lab / technical testing": "Independent price audit / comparison table",
    }
    for route in analysis["claim_routes"]:
        route["best_source_types"] = [replacements.get(item, item) for item in route["best_source_types"]]
    for bucket in analysis["who_needs_to_say_it"].values():
        if not isinstance(bucket, list):
            continue
        for source in bucket:
            if isinstance(source, dict) and source.get("source_type") in replacements:
                source["source_type"] = replacements[source["source_type"]]
    for source in analysis["third_party_story_map"]["sources"]:
        source["source_type"] = replacements.get(source.get("source_type"), source.get("source_type"))
    for source in analysis["current_web_reality"]:
        source["source_type"] = replacements.get(source.get("source_type"), source.get("source_type"))


def sanitize_unverified_source_names(
    analysis: dict[str, Any],
    verified_sources: list[dict[str, Any]],
) -> None:
    verified_names = {
        source["title"].strip().lower()
        for source in verified_sources
        if source.get("title")
    }
    targets = analysis["who_needs_to_say_it"]["source_targets_only"]
    target_types = []
    deduped_targets = []
    seen_types = set()
    for source in targets:
        source_type = source["source_type"]
        if source_type in seen_types:
            continue
        seen_types.add(source_type)
        source["source_name"] = source_type
        source["url"] = ""
        source["discovery_path_or_query"] = "Source target only; not returned by web search."
        source["current_stance"] = "Not scanned"
        source["what_it_currently_says"] = "Not scanned in this run."
        deduped_targets.append(source)
        target_types.append(source_type)
    analysis["who_needs_to_say_it"]["source_targets_only"] = deduped_targets

    for source in analysis["third_party_story_map"]["sources"]:
        if source.get("source_name", "").strip().lower() not in verified_names and not source.get("url"):
            source["source_name"] = source["source_type"]

    for source in analysis["validation_target_map"]:
        if source.get("source_name_or_type", "").strip().lower() not in verified_names:
            source["source_name_or_type"] = "Relevant validation source type"

    for index, item in enumerate(analysis["contact_priority_list"]):
        name = item["source_or_source_type"].strip()
        if name.lower() not in verified_names and target_types:
            item["source_or_source_type"] = target_types[index % len(target_types)]

    for plan_item in analysis["source_validation_plan"]:
        verified_targets = [
            name
            for name in plan_item["target_examples_or_discovered_sources"]
            if name.strip().lower() in verified_names
        ]
        plan_item["target_examples_or_discovered_sources"] = (
            verified_targets or [plan_item["source_role"]]
        )

    verified_titles = [
        source["title"]
        for source in verified_sources
        if source.get("title") and source.get("citation_backed")
    ]
    citation_version = analysis["argument_to_seed"]["owned_site_citation_version"]
    if not any(title.lower() in citation_version.lower() for title in verified_titles):
        analysis["argument_to_seed"]["owned_site_citation_version"] = (
            "According to [third-party comparison source, date], "
            + analysis["argument_to_seed"]["source_ready_argument"]
        )


def sanitize_unverified_currentness(
    analysis: dict[str, Any],
    source_discovery: dict[str, Any],
) -> None:
    scanned_sources = [
        source["title"]
        for source in source_discovery["sources"]
        if source.get("citation_backed")
    ]
    if scanned_sources:
        analysis["proof_packet"]["currentness"] = [
            f"Web discovery run on {date.today().isoformat()}.",
            "Recheck pricing, promotions, eligibility, availability, and competitor data before publication.",
            "Sources summarized in this run: " + ", ".join(scanned_sources[:10]),
        ]
    else:
        analysis["proof_packet"]["currentness"] = [
            "Web search ran, but no usable sources were discovered or scanned; currentness remains unverified."
        ]

    analysis["proof_packet"]["sourceable_facts"] = [
        fact if fact.lower().startswith("verify") else f"Verify: {fact}"
        for fact in analysis["proof_packet"]["sourceable_facts"]
    ]


def normalize_source_discovery(
    analysis: dict[str, Any],
    inputs: dict[str, Any],
    citations: list[dict[str, str]],
    web_activity: dict[str, Any],
) -> dict[str, Any]:
    research_mode = inputs["research_mode"]
    web_enabled = (
        not LIVE_WEB_DISABLED
        and research_mode in {"Web-grounded", "Evidence + web"}
    )
    web_queries = web_activity["queries"]
    discovered_web_sources = web_activity["sources"]
    citation_urls = {normalize_url(source["url"]) for source in citations if source.get("url")}
    discovered_urls = {
        normalize_url(source["url"]) for source in discovered_web_sources if source.get("url")
    }
    source_map = analysis["who_needs_to_say_it"]
    normalize_impact_labels(analysis)
    normalize_price_source_types(analysis, inputs["claim"])

    if not web_enabled:
        target_candidates = (
            source_map["discovered_sources"]
            + source_map["scanned_sources"]
            + source_map["source_targets_only"]
        )
        source_map["discovered_sources"] = []
        source_map["scanned_sources"] = []
        source_map["source_targets_only"] = []
        seen_targets = set()
        for source in target_candidates:
            key = (source["source_name"].lower(), source["source_type"])
            if key in seen_targets:
                continue
            seen_targets.add(key)
            source["source_status"] = "Source target only — not found/scanned in this run"
            source["current_stance"] = "Not checked"
            source["what_it_currently_says"] = "Not scanned in this run."
            source["url"] = ""
            source_map["source_targets_only"].append(source)
        analysis["current_web_reality"] = []
        for source in analysis["validation_target_map"]:
            source["source_status_badge"] = "Source target only"
        summary = analysis["source_discovery_summary"]
        summary["queries_run"] = []
        summary["sources_discovered"] = 0
        summary["sources_scanned"] = 0
        summary["source_mix"] = []
        summary["key_gap"] = "No current web evidence was checked."
        summary["highest_leverage_source_opportunity"] = (
            "Run web-grounded discovery to identify the sources AI currently synthesizes for this recommendation."
        )
        summary["evidence_layer_summary"] = "No web sources were searched or scanned in this run."
        analysis["confidence_source_coverage"] = {
            "status": "Not checked",
            "note": "No web queries or source scans were run.",
        }
        analysis["third_party_story_map"]["coverage_note"] = "No web queries run. Source Target Map only."
        for source in analysis["third_party_story_map"]["sources"]:
            source["scan_status"] = "source target only"
            source["current_stance"] = "not scanned"
            source["what_it_currently_says"] = "Not scanned."
            source["url"] = ""
        analysis["argument_to_seed"]["owned_site_citation_version"] = (
            "According to [third-party comparison source, date], "
            + analysis["argument_to_seed"]["source_ready_argument"]
        )
        return analysis

    all_sources = (
        source_map["discovered_sources"]
        + source_map["scanned_sources"]
        + source_map["source_targets_only"]
    )
    modeled_by_url = {
        normalize_url(source.get("url", "")): source
        for source in all_sources
        if source.get("url")
    }
    for raw_source in discovered_web_sources:
        url = normalize_url(raw_source["url"])
        if url in modeled_by_url:
            continue
        all_sources.append(
            {
                "source_name": raw_source["title"],
                "url": raw_source["url"],
                "discovery_path_or_query": ", ".join(
                    raw_source.get("discovery_queries", [])
                ) or "OpenAI web search",
                "source_type": "Other",
                "source_status": "Discovered, not scanned",
                "why_relevant_to_recommendation_ecosystem": (
                    "This source appeared in the web-search evidence layer for the recommendation query."
                ),
                "what_it_currently_says": " ".join(raw_source.get("summaries", []))
                or "Discovered through search but not summarized in the final response.",
                "current_stance": (
                    "Partially supports"
                    if raw_source.get("citation_backed") and raw_source.get("summaries")
                    else "Not scanned"
                ),
                "what_it_would_need_to_say": analysis["argument_to_seed"]["source_ready_argument"],
                "likely_impact_if_echoed_upgraded_claim": "Mention-worthy",
                "source_role": "Discovered web evidence source",
                "can_brand_influence": "Limited",
                "contact_action_priority": "Low",
                "recommended_brand_action": "Review this source and assess whether it is a relevant validation or outreach target.",
            }
        )

    scanned = []
    discovered = []
    targets = []
    seen_sources = set()
    for source in all_sources:
        key = (normalize_url(source["url"]), source["source_name"].lower())
        if key in seen_sources:
            continue
        seen_sources.add(key)
        if normalize_url(source["url"]) in citation_urls:
            source["source_status"] = "Scanned"
            if source["current_stance"] == "not scanned":
                source["current_stance"] = "Not scanned"
            scanned.append(source)
        elif normalize_url(source["url"]) in discovered_urls:
            source["source_status"] = "Discovered, not scanned"
            source["current_stance"] = "Not scanned"
            source["what_it_currently_says"] = "Discovered through search but not summarized in the final response."
            discovered.append(source)
        else:
            source["source_status"] = "Source target only — not found/scanned in this run"
            source["current_stance"] = "Not scanned"
            source["what_it_currently_says"] = "Not scanned in this run."
            source["url"] = ""
            targets.append(source)

    source_map["discovered_sources"] = discovered
    source_map["scanned_sources"] = list(scanned)
    source_map["source_targets_only"] = targets

    summary = analysis["source_discovery_summary"]
    summary["queries_run"] = web_queries
    summary["sources_discovered"] = len(discovered) + len(scanned)
    summary["sources_scanned"] = len(scanned)
    summary["source_mix"] = sorted({source["source_type"] for source in scanned})
    source_names = ", ".join(source["source_name"] for source in scanned)
    if source_names:
        summary["evidence_layer_summary"] = (
            "Based on the discovered evidence layer, the sources most likely to influence "
            f"an AI answer in this run are: {source_names}."
        )
    elif discovered:
        summary["evidence_layer_summary"] = (
            f"Web search discovered {len(discovered)} source(s), but none were "
            "citation-backed or summarized as scanned in this run."
        )
    elif web_activity["web_search_call_count"] > 0:
        summary["evidence_layer_summary"] = (
            "Web search ran, but no usable sources were discovered or scanned."
        )
    else:
        summary["evidence_layer_summary"] = (
            "Web search was attempted, but no web search call completed."
        )
    coverage_status, coverage_note = evidence_coverage(
        len(scanned),
        summary["source_mix"],
        web_activity["web_search_call_count"],
    )
    if coverage_status != "Robust":
        coverage_note += " For stronger confidence, scan 8-10 surfaces across editorial, review aggregator, competitor, and community evidence."
    analysis["confidence_source_coverage"] = {
        "status": coverage_status,
        "note": coverage_note,
    }
    analysis["third_party_story_map"]["coverage_note"] = coverage_note
    analysis["_web_debug"] = {
        "web_tool_enabled": web_enabled,
        "tool_choice": "required" if web_enabled else "none",
        "web_search_calls_detected": web_activity["web_search_call_count"],
        "search_queries_detected": web_queries,
        "sources_discovered": len(discovered) + len(scanned),
        "sources_scanned_opened": max(
            len(scanned),
            web_activity["opened_page_count"],
        ),
    }

    analysis["current_web_reality"] = [
        source
        for source in analysis["current_web_reality"]
        if normalize_url(source.get("url", "")) in citation_urls
    ]
    for source in analysis["validation_target_map"]:
        source["source_status_badge"] = "Source target only"
    if not citations:
        analysis["argument_to_seed"]["owned_site_citation_version"] = (
            "According to [third-party comparison source, date], "
            + analysis["argument_to_seed"]["source_ready_argument"]
        )
    for source in analysis["third_party_story_map"]["sources"]:
        if normalize_url(source.get("url", "")) in citation_urls:
            source["scan_status"] = "found and scanned"
        else:
            source["scan_status"] = "source target only"
            source["current_stance"] = "not scanned"
            source["what_it_currently_says"] = "Not scanned."
            source["url"] = ""
    sanitize_unverified_source_names(analysis, discovered_web_sources)
    return analysis


def build_openai_request(
    inputs: dict[str, Any],
    evidence_text: str,
    source_discovery: dict[str, Any],
) -> dict[str, Any]:
    return dict(
        model=MODEL,
        input=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {
                "role": "user",
                "content": build_user_prompt(inputs, evidence_text, source_discovery),
            },
        ],
        temperature=0.2,
        max_output_tokens=10000,
        text={"format": {"type": "json_schema", **ANALYSIS_SCHEMA}},
    )


def analyze_claim(api_key: str, inputs: dict[str, Any], evidence_text: str) -> dict[str, Any]:
    client = OpenAI(api_key=api_key)
    web_enabled = (
        not LIVE_WEB_DISABLED
        and inputs["research_mode"] in {"Web-grounded", "Evidence + web"}
    )
    source_discovery = (
        run_source_discovery(client, inputs)
        if web_enabled
        else {
            "search_intent": "No web discovery performed.",
            "planned_queries": [],
            "queries": [],
            "sources": [],
            "summaries": [],
            "citations": [],
            "web_search_call_count": 0,
            "opened_page_count": 0,
            "errors": [],
        }
    )
    request_args = build_openai_request(inputs, evidence_text, source_discovery)

    response = client.responses.create(**request_args)

    analysis = extract_json_from_response(response)
    response_citations = source_discovery["citations"]
    web_activity = {
        "web_search_call_count": source_discovery["web_search_call_count"],
        "opened_page_count": source_discovery["opened_page_count"],
        "queries": source_discovery["queries"],
        "sources": source_discovery["sources"],
    }
    analysis["citations"] = response_citations
    analysis = normalize_source_discovery(
        analysis,
        inputs,
        response_citations,
        web_activity,
    )
    if web_enabled:
        sanitize_unverified_currentness(analysis, source_discovery)
        analysis["_web_debug"]["discovery_errors"] = source_discovery["errors"]
        analysis["_web_debug"]["analysis_call_used_web_tool"] = False
    if not web_enabled:
        analysis["_web_debug"] = {
            "web_tool_enabled": False,
            "tool_choice": "none",
            "web_search_calls_detected": 0,
            "search_queries_detected": [],
            "sources_discovered": 0,
            "sources_scanned_opened": 0,
            "discovery_errors": [],
            "analysis_call_used_web_tool": False,
        }
    analysis["_input_context"] = dict(inputs)
    analysis["_analysis_version"] = ANALYSIS_VERSION
    return analysis


def card(title: str, body: str) -> None:
    st.markdown(f"#### {title}")
    st.markdown(body)


def list_cards(items: list[dict[str, str]], title_key: str, body_keys: list[str]) -> None:
    for item in items:
        with st.container(border=True):
            st.markdown(f"**{item.get(title_key, '')}**")
            for key in body_keys:
                if item.get(key):
                    label = key.replace("_", " ").capitalize()
                    st.markdown(f"**{label}:** {item[key]}")


def scorecard_grid(cards: list[tuple[str, str]]) -> None:
    card_html = "".join(
        (
            '<div class="reclaimer-scorecard">'
            f'<div class="reclaimer-scorecard-label">{html.escape(label)}</div>'
            f'<div class="reclaimer-scorecard-value">{html.escape(value)}</div>'
            "</div>"
        )
        for label, value in cards
    )
    st.markdown(f'<div class="reclaimer-scorecard-grid">{card_html}</div>', unsafe_allow_html=True)


def evidence_block_as_yaml(block: dict[str, Any]) -> str:
    lines = []
    for key, value in block.items():
        label = key.replace("_", " ")
        if isinstance(value, list):
            lines.append(f"{label}:")
            lines.extend([f"  - {entry}" for entry in value])
        else:
            lines.append(f"{label}: {value}")
    return "\n".join(lines)


def export_markdown(analysis: dict[str, Any]) -> dict[str, str]:
    structure = analysis["claim_structure_status"]
    input_context = analysis.get("_input_context", {})
    web_debug = analysis.get("_web_debug", {})
    input_context_md = "\n".join(
        [
            f"- **Brand / product:** {input_context.get('brand', '[Not provided]')}",
            f"- **Category:** {input_context.get('category', '[Not provided]')}",
            f"- **Claim:** {input_context.get('claim', '[Not provided]')}",
            f"- **Known facts:** {input_context.get('known_facts') or '[Not provided]'}",
            f"- **Competitors:** {input_context.get('competitors') or '[Not provided]'}",
            f"- **Research mode:** {input_context.get('research_mode', '[Not provided]')}",
        ]
    )
    claim_routes_md = "\n".join(
        "- **{name}**: {claim} Lane: {lane} Proof: {proof} Sources: {sources} Impact: {impact} Risks: {risks}".format(
            name=item["route_name"],
            claim=item["stronger_source_ready_claim"],
            lane=item["user_intent_or_recommendation_lane"],
            proof=", ".join(item["proof_needed"]),
            sources=", ".join(item["best_source_types"]),
            impact=item["likely_impact_if_validated"],
            risks=", ".join(item["main_tradeoffs_or_risks"]),
        )
        for item in analysis["claim_routes"]
    )
    discovery = analysis["source_discovery_summary"]
    who_sources = "\n".join(
        "- **{name}** ({status}; {source_type}): Query/path: {path}. Current: {current} Stance: {stance}. Needs: {needed}. AI relevance: {care}. Impact: {impact}. Influence: {influence}. Priority: {priority}. Action: {action} {url}".format(
            name=item["source_name"],
            status=item["source_status"],
            source_type=item["source_type"],
            path=item["discovery_path_or_query"],
            current=item["what_it_currently_says"],
            stance=item["current_stance"],
            needed=item["what_it_would_need_to_say"],
            care=item["why_relevant_to_recommendation_ecosystem"],
            impact=item["likely_impact_if_echoed_upgraded_claim"],
            influence=item["can_brand_influence"],
            priority=item["contact_action_priority"],
            action=item["recommended_brand_action"],
            url=item["url"],
        )
        for group in [
            analysis["who_needs_to_say_it"]["discovered_sources"],
            analysis["who_needs_to_say_it"]["scanned_sources"],
            analysis["who_needs_to_say_it"]["source_targets_only"],
        ]
        for item in group
    )
    contact_priorities = "\n".join(
        f"{item['priority_rank']}. **{item['source_or_source_type']}**: "
        f"{item['why_it_matters']} Give/do: {item['what_to_give_or_do']} "
        f"Expected impact: {item['expected_impact_if_successful']}"
        for item in sorted(analysis["contact_priority_list"], key=lambda entry: entry["priority_rank"])
    )
    route_ladders = "\n".join(
        f"### {route['route_name']}\n"
        + "\n".join(
            f"- **{stage['echo_stage']} → {stage['impact']}**: {stage['why']}"
            for stage in route["stages"]
        )
        for route in analysis["source_echo_impact_ladder"]
    )
    confidence = analysis["confidence_source_coverage"]
    argument = analysis["argument_to_seed"]
    proof = analysis["proof_packet"]
    proof_md = "\n".join(
        [
            f"**Metrics:** {', '.join(proof['metrics'])}",
            f"**Comparison set:** {', '.join(proof['comparison_set'])}",
            f"**Sourceable facts:** {', '.join(proof['sourceable_facts'])}",
            f"**Caveats:** {', '.join(proof['caveats'])}",
            f"**Methodology:** {', '.join(proof['methodology'])}",
            f"**Currentness:** {', '.join(proof['currentness'])}",
            f"**Artifacts needed:** {', '.join(proof['artifacts_needed'])}",
        ]
    )
    source_echo_md = "\n".join(
        "- **{source_type}**: {reading} Authority: {authority} Evidence: {evidence} Impact: {impact}".format(
            source_type=item["source_type"],
            reading=item["how_argument_would_read"],
            authority=item["authority_upgrade"],
            evidence=item["surrounding_evidence_required"],
            impact=item["likely_recommendation_impact"],
        )
        for item in analysis["source_echo_simulation"]
    )
    intent = analysis["intent_match"]
    general_impact = analysis["general_category_impact"]
    intent_impact = analysis["intent_specific_impact"]
    criteria = analysis["decision_criteria_detected"]
    criterion_support = "\n".join(
        f"- **{item['criterion']}** ({item['support_type']}): {item['evidence_from_claim']}"
        for item in criteria["criterion_support"]
    )
    story_sources = "\n".join(
        "- **{name}** ({group}; {scan_status}; {stance}): {current} Needs: {needed} AI relevance: {care} Action: {action} {url}".format(
            name=item["source_name"],
            group=item["source_group"],
            scan_status=item["scan_status"],
            stance=item["current_stance"],
            current=item["what_it_currently_says"],
            needed=item["what_it_would_need_to_say"],
            care=item["why_an_ai_agent_would_care"],
            action=item["recommended_brand_action"],
            url=item["url"],
        )
        for item in analysis["third_party_story_map"]["sources"]
    )
    current_web = "\n".join(
        "- **{name}** ({source_type}) [{badge}]: {status}. URL: {url}. Current: {current} Supports claim: {supports}. Notes: {notes}".format(
            name=item["source_name"],
            source_type=item["source_type"],
            badge=item["source_status_badge"],
            status=item["status"],
            url=item["url"],
            current=item["what_it_currently_says"],
            supports=item["supports_claim"],
            notes=item["notes_or_caveats"],
        )
        for item in analysis["current_web_reality"]
    ) or "No sources were checked in this run."
    target_map = "\n".join(
        "- **{name}** [{badge}]: {care} Evidence needed: {evidence} Brand influence: {influence} Action: {action}".format(
            name=item["source_name_or_type"],
            badge=item["source_status_badge"],
            care=item["why_an_ai_agent_would_care"],
            evidence=item["evidence_that_would_need_to_appear"],
            influence=item["can_brand_ethically_influence_it"],
            action=item["recommended_action"],
        )
        for item in analysis["validation_target_map"]
    )

    def echo_md(items: list[dict[str, str]]) -> str:
        return "\n".join(
            "- **{stage}** [{badge}]: {impact}. Behavior: {behavior} Why: {why} Improvement needed: {improvement} Source validation: {source}".format(
            stage=item["stage"],
            badge=item["source_status_badge"],
            impact=item["impact_label"],
            behavior=item["likely_agent_behavior"],
            why=item["why"],
            improvement=item["structural_improvement_needed"],
            source=item["source_validation_needed"],
        )
            for item in items
        )

    exact_echo = echo_md(analysis["counterfactual_source_echo"]["exact_claim_as_written"])
    upgraded_echo = echo_md(analysis["counterfactual_source_echo"]["recommended_upgraded_version"])
    elements = "\n".join(
        "- **{element}** ({classification}): Missing: {missing} Why it matters: {why}".format(
            element=item["element"],
            classification=item["classification"],
            missing=item["what_is_missing"],
            why=item["why_it_matters"],
        )
        for item in analysis["claim_elements"]
    )
    audit = analysis["math_logic_audit"]
    audit_md = f"""Numbers detected: {audit["numbers_detected"]}

Percentages: {", ".join(audit["percentages"]) or "None detected"}

Savings / price claims: {", ".join(audit["savings_or_price_claims"]) or "None detected"}

X-times claims: {", ".join(audit["times_claims"]) or "None detected"}

Smaller / larger claims: {", ".join(audit["size_change_claims"]) or "None detected"}

Questions to resolve:
{chr(10).join(f"- {item}" for item in audit["math_questions_to_resolve"]) or "- None detected"}

Wording accuracy risk: {audit["wording_accuracy_risk"]}
"""
    stronger = analysis["stronger_agent_ready_argument"]
    stronger_md = "\n".join(
        f"- **{item['move']}**: {item['why_it_makes_the_claim_more_agent_usable']}"
        for item in stronger["path_from_marketing_to_decision_proof"]
    )
    source_plan = "\n".join(
        "- **{role}**: {care} Targets: {targets} Needs to say: {needs} Evidence: {evidence} Brand influence: {influence} Action: {action}".format(
            role=item["source_role"],
            care=item["why_an_ai_agent_would_care"],
            targets=", ".join(item["target_examples_or_discovered_sources"]),
            needs=item["what_sources_need_to_say"],
            evidence=item["exact_evidence_that_should_appear"],
            influence=item["can_brand_ethically_influence_it"],
            action=item["suggested_action"],
        )
        for item in analysis["source_validation_plan"]
    )
    missing = "\n".join(
        f"- **{item['piece']}**: {item['why_it_matters']} Needed: {item['example_needed']}"
        for item in analysis["missing_evidence"]
    )
    rewrites = analysis["safer_claim_rewrites"]
    plan = "\n".join(
        f"- **{item['surface']}**: {item['what_to_publish_or_collect']} Why it matters: {item['why_an_ai_agent_would_care']}"
        for item in analysis["triangulation_plan"]
    )
    red_team = "\n".join(
        f"- **{item['risk']}**: {item['how_it_could_be_attacked_or_ignored']} Mitigation: {item['how_to_neutralize']}"
        for item in analysis["red_team"]
    )
    evidence = evidence_block_as_yaml(analysis["agent_readable_evidence_block"])

    full = f"""# Agent Evidence Plan

## Input Context
{input_context_md}

## Web Search Debug
- Web tool enabled: {str(web_debug.get("web_tool_enabled", False)).lower()}
- Tool choice: {web_debug.get("tool_choice", "none")}
- Web search calls detected: {web_debug.get("web_search_calls_detected", 0)}
- Search queries detected: {", ".join(web_debug.get("search_queries_detected", [])) or "None"}
- Sources discovered: {web_debug.get("sources_discovered", 0)}
- Sources scanned/opened: {web_debug.get("sources_scanned_opened", 0)}
- Main planner call used web tool: {str(web_debug.get("analysis_call_used_web_tool", False)).lower()}
- Discovery errors: {"; ".join(web_debug.get("discovery_errors", [])) or "None"}

## Top Verdict / Scorecard
- Claim as Written: {structure["status"]}
- Broad Category Impact: {general_impact["status"]}
- Impact for This User Need: {intent_impact["status"]}
- Simulated Agent Impact: {analysis["recommendation_impact"]["status"]}
- Evidence Coverage: {confidence["status"]} — {confidence["note"]}

## Simulated Recommendation
{analysis["simulated_recommendation"]}

## Claim Routes
{claim_routes_md}

## Recommended Route
**{analysis["recommended_route"]["route_name"]}**

{analysis["recommended_route"]["why_this_route"]}

Recommendation potential: {analysis["recommended_route"]["recommendation_potential"]}

## Intent Match
**Inferred user intent:** {intent["inferred_user_intent"]}

**Job to be done:** {intent["job_to_be_done"]}

{intent["how_claim_maps_to_intent"]}

## General Category Impact
**{general_impact["status"]}**: {general_impact["reason"]}

## Intent-Specific Impact
**{intent_impact["status"]}** for {intent_impact["intent"]}: {intent_impact["reason"]}

Tradeoffs: {intent_impact["tradeoff_note"]}

## Decision Criteria Detected
**Primary:** {criteria["primary_decision_criterion"]}

**Secondary:** {", ".join(criteria["secondary_decision_criteria"])}

{criterion_support}

**Tradeoffs:** {", ".join(criteria["tradeoffs"])}

**Missing proof:** {", ".join(criteria["missing_proof"])}

## Argument to Seed
### Source-Ready Argument
{argument["source_ready_argument"]}

### Owned-Site Citation Version
{argument["owned_site_citation_version"]}

## Proof Packet
{proof_md}

## Agent Source Map: Who Shapes the Answer
**Search intent:** {discovery["search_intent_used"]}

**Queries run:** {", ".join(discovery["queries_run"]) or "None"}

**Sources discovered/scanned:** {discovery["sources_discovered"]}/{discovery["sources_scanned"]}

**Source mix:** {", ".join(discovery["source_mix"]) or "None"}

**Key gap:** {discovery["key_gap"]}

**Highest-leverage source opportunity:** {discovery["highest_leverage_source_opportunity"]}

{discovery["evidence_layer_summary"]}

{who_sources}

### Contact Priority List
{contact_priorities}

## Source Echo Impact Ladder
{route_ladders}

## Third-Party Story Map
{analysis["third_party_story_map"]["coverage_note"]}

{story_sources}

## Source Echo Simulation
{source_echo_md}

## Counterfactual Source Echo
### A. Exact Claim As Written
{exact_echo}

### B. Source-Ready Argument
{upgraded_echo}

## Current Web Reality
{current_web}

## Validation Target Map
{target_map}

## Useful vs. Fluff / Needs Clarification
{elements}

## Math / Logic Audit
{audit_md}

## Stronger Agent-Ready Argument
{stronger["structured_rewrite"]}

### Path From Marketing To Decision Proof
{stronger_md}

### Decision Proof Template
{stronger["decision_proof_template"]}

## Supplemental Source Surface Plan
{source_plan}

## Verdict
{analysis["claim_verdict"]["summary"]}

## Claim Type
{", ".join(analysis["claim_type"])}

## What The AI Would Say Today
{analysis["what_ai_would_say_today"]}

## Missing Evidence
{missing}

## Safer Claim Rewrites
- Safe marketing-language version: {rewrites["safe_marketing_language"]}
- More substantiated version: {rewrites["more_substantiated_version"]}
- Recommendation-ready version: {rewrites["recommendation_ready_version_assuming_evidence_exists"]}

## Triangulation Plan
{plan}

## Red Team
{red_team}

## Recommendation Readiness
**{analysis["recommendation_readiness"]["label"]}**: {analysis["recommendation_readiness"]["rationale"]}

## Agent-Readable Evidence Block
```yaml
{evidence}
```
"""
    return {
        "Full agent evidence plan.md": full,
        "Missing evidence checklist.md": f"# Missing Evidence Checklist\n\n{missing}\n",
        "Safer claim rewrites.md": (
            "# Safer Claim Rewrites\n\n"
            f"## Safe Marketing Language\n{rewrites['safe_marketing_language']}\n\n"
            f"## More Substantiated Version\n{rewrites['more_substantiated_version']}\n\n"
            "## Recommendation-Ready Version\n"
            f"{rewrites['recommendation_ready_version_assuming_evidence_exists']}\n"
        ),
        "Agent-readable evidence block.yaml": evidence,
        "Triangulation plan.md": f"# Triangulation Plan\n\n{plan}\n\n## Supplemental Target Surface Plan\n\n{source_plan}\n",
        "Red team report.md": f"# Red Team Report\n\n{red_team}\n",
    }


st.set_page_config(page_title=APP_TITLE, layout="wide")
st.markdown(
    """
    <style>
    .reclaimer-scorecard-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(210px, 1fr));
        gap: 14px;
        margin: 0.25rem 0 0.75rem;
    }
    .reclaimer-scorecard {
        min-height: 132px;
        padding: 18px 20px;
        border: 1px solid rgba(128, 128, 128, 0.24);
        border: 1px solid color-mix(in srgb, currentColor 18%, transparent);
        border-radius: 8px;
        color: inherit;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
    }
    .reclaimer-scorecard-label {
        color: inherit;
        opacity: 0.7;
        font-size: 0.9rem;
        line-height: 1.3;
        font-weight: 600;
    }
    .reclaimer-scorecard-value {
        color: inherit;
        font-size: 1.65rem;
        line-height: 1.15;
        font-weight: 500;
        overflow-wrap: anywhere;
        margin-top: 14px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)
st.title(APP_TITLE)
st.caption(UNIVERSAL_RECOMMENDATION_QUESTION)

configured_api_key, _api_key_source = load_api_key()
with st.sidebar:
    st.header("AI Settings")
    if LIVE_WEB_DISABLED:
        st.warning("Demo fallback mode: live web search is disabled.")
    if configured_api_key:
        st.success("API key configured for demo.")
        active_api_key = configured_api_key
    else:
        active_api_key = st.text_input(
            "OpenAI API key",
            type="password",
            help="Used only for this session. The key is not printed, logged, or exported.",
        )
    st.caption("Default model: " + MODEL)

    preset_name = st.selectbox("Sample claims", list(SAMPLE_CLAIMS.keys()))
    selected_preset = SAMPLE_CLAIMS[preset_name]

if "analysis" not in st.session_state:
    st.session_state.analysis = None

with st.form("claim_form"):
    col_a, col_b = st.columns([1, 1])
    with col_a:
        brand = st.text_input(
            "Brand / product",
            value=selected_preset.get("brand", ""),
            placeholder="Roadster 2.0",
        )
        category = st.text_input(
            "Category",
            value=selected_preset.get("category", ""),
            placeholder="Running shoes",
        )
        claim = st.text_area(
            "Claim",
            value=selected_preset.get("claim", ""),
            placeholder="The most durable running shoe under $150.",
            height=110,
        )
    with col_b:
        known_facts = st.text_area(
            "Known facts (optional)",
            placeholder="Any proof, caveats, test results, pricing, market, dates, or eligibility details you already know.",
            height=110,
        )
        competitors = st.text_area(
            "Competitors (optional)",
            placeholder="Known competitors or comparison set, if any.",
            height=110,
        )
        default_research_mode = (
            "Web-grounded"
            if active_api_key and "Web-grounded" in RESEARCH_MODES
            else "No web"
        )
        research_mode = st.selectbox(
            "Research mode",
            RESEARCH_MODES,
            index=RESEARCH_MODES.index(default_research_mode),
        )

    uploaded_files = st.file_uploader(
        "Uploaded evidence files (optional)",
        type=["txt", "md", "pdf", "docx", "csv"],
        accept_multiple_files=True,
    )
    submitted = st.form_submit_button("Plan agent evidence", type="primary")

if submitted:
    if not active_api_key:
        st.error(
            "AI mode requires an OpenAI API key. Configure OPENAI_API_KEY in the "
            "deployment environment, Streamlit secrets, or the sidebar."
        )
    elif not brand.strip() or not category.strip() or not claim.strip():
        st.error("Add a brand / product, category, and claim before running the planner.")
    else:
        evidence_chunks = []
        for uploaded_file in uploaded_files:
            with st.spinner(f"Reading {uploaded_file.name}"):
                text = get_text_from_upload(uploaded_file)
                evidence_chunks.append(f"--- {uploaded_file.name} ---\n{text[:12000]}")
        evidence_text = "\n\n".join(evidence_chunks)
        inputs = {
            "brand": brand.strip(),
            "claim": claim.strip(),
            "category": category.strip(),
            "known_facts": known_facts.strip(),
            "competitors": competitors.strip(),
            "research_mode": research_mode,
        }
        with st.spinner("Planning the agent-ready evidence path..."):
            try:
                st.session_state.analysis = analyze_claim(active_api_key, inputs, evidence_text)
            except Exception as exc:
                st.session_state.analysis = None
                st.error(f"Analysis failed: {exc}")

analysis = st.session_state.analysis

required_analysis_sections = {
    "claim_routes",
    "recommended_route",
    "source_discovery_summary",
    "who_needs_to_say_it",
    "source_echo_impact_ladder",
    "contact_priority_list",
    "_web_debug",
    "argument_to_seed",
    "proof_packet",
    "source_echo_simulation",
    "confidence_source_coverage",
    "_input_context",
    "_analysis_version",
}
if analysis and (
    not required_analysis_sections.issubset(analysis)
    or analysis.get("_analysis_version") != ANALYSIS_VERSION
):
    st.session_state.analysis = None
    analysis = None
    st.info("The planner output format was updated. Run the claim again to generate the new source-echo analysis.")

if not analysis:
    st.info("Enter a brand / product, category, and claim. No web is used unless you select a web-grounded research mode.")
    st.stop()

impact = analysis["recommendation_impact"]
structure = analysis["claim_structure_status"]
general_impact = analysis["general_category_impact"]
intent_impact = analysis["intent_specific_impact"]
confidence = analysis["confidence_source_coverage"]
intent = analysis["intent_match"]
criteria = analysis["decision_criteria_detected"]
story_map = analysis["third_party_story_map"]
input_context = analysis["_input_context"]
web_debug = analysis["_web_debug"]

st.markdown(
    f"**{input_context['brand']}** · {input_context['category']} · "
    f"{input_context['research_mode']}"
)
st.caption(f"Claim: {input_context['claim']}")
if input_context.get("known_facts"):
    st.caption(f"Known facts: {input_context['known_facts']}")
if input_context.get("competitors"):
    st.caption(f"Competitors: {input_context['competitors']}")

st.subheader("Top Verdict / Scorecard")
scorecard_grid(
    [
        ("Claim as Written", structure["status"]),
        ("Broad Category Impact", general_impact["status"]),
        ("Impact for This User Need", intent_impact["status"]),
        ("Simulated Agent Impact", impact["status"]),
        ("Evidence Coverage", confidence["status"]),
    ]
)
st.caption(confidence["note"])

with st.expander("Web Search Debug", expanded=input_context["research_mode"] in {"Web-grounded", "Evidence + web"}):
    debug_cols = st.columns(3)
    debug_cols[0].metric("Web tool enabled", str(web_debug["web_tool_enabled"]).lower())
    debug_cols[1].metric("Tool choice", web_debug["tool_choice"])
    debug_cols[2].metric("Web search calls", web_debug["web_search_calls_detected"])
    st.markdown(
        "**Search queries detected:** "
        + (", ".join(web_debug["search_queries_detected"]) or "None")
    )
    source_debug_cols = st.columns(2)
    source_debug_cols[0].metric("Sources discovered", web_debug["sources_discovered"])
    source_debug_cols[1].metric(
        "Sources scanned / opened",
        web_debug["sources_scanned_opened"],
    )
    st.markdown(
        f"**Main planner call used web tool:** "
        f"{str(web_debug.get('analysis_call_used_web_tool', False)).lower()}"
    )
    if web_debug.get("discovery_errors"):
        st.warning("Some discovery searches failed:\n\n" + "\n".join(web_debug["discovery_errors"]))

with st.container(border=True):
    st.subheader("Simulated Recommendation")
    st.markdown(analysis["simulated_recommendation"])

with st.container(border=True):
    st.subheader("Claim Routes")
    for route in analysis["claim_routes"]:
        st.markdown(f"**{route['route_name']}**")
        st.markdown(f"**What the claim is trying to become:** {route['what_current_claim_is_trying_to_become']}")
        st.code(route["stronger_source_ready_claim"], language="markdown")
        st.markdown(f"**Recommendation lane:** {route['user_intent_or_recommendation_lane']}")
        st.markdown("**Key decision criteria:** " + ", ".join(route["key_decision_criteria"]))
        st.markdown("**Proof needed:** " + ", ".join(route["proof_needed"]))
        st.markdown("**Best source types:** " + ", ".join(route["best_source_types"]))
        st.markdown(f"**Likely impact if validated:** {route['likely_impact_if_validated']}")
        st.markdown("**Tradeoffs / risks:** " + ", ".join(route["main_tradeoffs_or_risks"]))
        st.divider()

with st.container(border=True):
    st.subheader("Recommended Route")
    recommended_route = analysis["recommended_route"]
    st.markdown(f"**{recommended_route['route_name']}**")
    st.markdown(recommended_route["why_this_route"])
    st.markdown(f"**Recommendation potential:** {recommended_route['recommendation_potential']}")

with st.container(border=True):
    st.subheader("Argument to Seed")
    argument = analysis["argument_to_seed"]
    st.markdown("**Source-ready argument**")
    st.code(argument["source_ready_argument"], language="markdown")
    st.caption(argument["claim_payload_note"])
    st.markdown("**Owned-site citation version**")
    st.code(argument["owned_site_citation_version"], language="markdown")

with st.container(border=True):
    st.subheader("Proof Packet")
    proof = analysis["proof_packet"]
    proof_cols = st.columns(2)
    with proof_cols[0]:
        st.markdown("**Metrics**")
        st.write(proof["metrics"])
        st.markdown("**Comparison set**")
        st.write(proof["comparison_set"])
        st.markdown("**Sourceable facts**")
        st.write(proof["sourceable_facts"])
        st.markdown("**Caveats**")
        st.write(proof["caveats"])
    with proof_cols[1]:
        st.markdown("**Methodology**")
        st.write(proof["methodology"])
        st.markdown("**Currentness**")
        st.write(proof["currentness"])
        st.markdown("**Screenshots / tables / exports needed**")
        st.write(proof["artifacts_needed"])

with st.container(border=True):
    st.subheader("Agent Source Map: Who Shapes the Answer")
    st.caption(
        "These are the sources AI is likely to synthesize when answering this recommendation "
        "question — and what each source would need to say for the claim to move."
    )
    discovery = analysis["source_discovery_summary"]
    st.markdown("**Agent Source Summary**")
    if not discovery["queries_run"]:
        st.info("Source Target Map only. No web queries run.")
    st.markdown(f"**Search intent used:** {discovery['search_intent_used']}")
    st.markdown(f"**Queries run:** {', '.join(discovery['queries_run']) or 'No web queries run.'}")
    discovery_cols = st.columns(2)
    discovery_cols[0].metric("Sources Discovered", discovery["sources_discovered"])
    discovery_cols[1].metric("Sources Scanned", discovery["sources_scanned"])
    st.markdown(f"**Source mix:** {', '.join(discovery['source_mix']) or 'None'}")
    st.markdown(f"**Key gap:** {discovery['key_gap']}")
    st.markdown(f"**Highest-leverage source opportunity:** {discovery['highest_leverage_source_opportunity']}")
    st.markdown(f"**Evidence layer:** {discovery['evidence_layer_summary']}")

    source_groups = analysis["who_needs_to_say_it"]
    for title, key in [
        ("Discovered Sources", "discovered_sources"),
        ("Scanned Sources", "scanned_sources"),
        ("Source Targets Only", "source_targets_only"),
    ]:
        sources = source_groups[key]
        if not sources:
            continue
        st.markdown(f"**{title}**")
        for source in sources:
            st.markdown(f"**{source['source_name']}**")
            st.caption(f"{source['source_status']} | {source['source_type']}")
            if source["url"]:
                st.markdown(f"[{source['url']}]({source['url']})")
            st.markdown(f"**Discovery path:** {source['discovery_path_or_query']}")
            st.markdown(f"**Current stance:** {source['current_stance']}")
            st.markdown(f"**What it currently says:** {source['what_it_currently_says']}")
            st.markdown(f"**What it would need to say:** {source['what_it_would_need_to_say']}")
            st.markdown(f"**Impact if it echoed the upgraded claim:** {source['likely_impact_if_echoed_upgraded_claim']}")
            st.markdown(f"**Source role:** {source['source_role']}")
            st.markdown(f"**Why AI cares:** {source['why_relevant_to_recommendation_ecosystem']}")
            st.markdown(f"**Can brand influence it?** {source['can_brand_influence']}")
            st.markdown(f"**Contact / action priority:** {source['contact_action_priority']}")
            st.markdown(f"**Brand action:** {source['recommended_brand_action']}")
            st.divider()

    if source_groups["missing_source_types"]:
        st.markdown("**Missing Source Types**")
        for gap in source_groups["missing_source_types"]:
            st.markdown(
                f"- **{gap['source_type']}**: {gap['why_missing_evidence_matters']} "
                f"Action: {gap['recommended_search_or_action']}"
            )

    st.markdown("**Contact Priority List**")
    for item in sorted(analysis["contact_priority_list"], key=lambda entry: entry["priority_rank"]):
        st.markdown(f"**{item['priority_rank']}. {item['source_or_source_type']}**")
        st.markdown(f"**Why it matters:** {item['why_it_matters']}")
        st.markdown(f"**What to give them / do:** {item['what_to_give_or_do']}")
        st.markdown(f"**Expected impact:** {item['expected_impact_if_successful']}")

with st.container(border=True):
    st.subheader("Source Echo Impact Ladder")
    for route in analysis["source_echo_impact_ladder"]:
        st.markdown(f"**{route['route_name']}**")
        for stage in route["stages"]:
            st.markdown(f"- **{stage['echo_stage']} → {stage['impact']}**: {stage['why']}")
        st.divider()

if analysis.get("citations"):
    st.subheader("Sources")
    for source in analysis["citations"]:
        st.markdown(f"- [{source['title']}]({source['url']}) - {source['note']}")

st.divider()
st.subheader("Secondary Detail")

tabs = st.tabs(
    [
        "Counterfactual Source Echo",
        "Missing Evidence",
        "AI-Safe Language",
        "Triangulation Plan",
        "Red Team",
        "Evidence Package",
        "Export",
    ]
)

with tabs[0]:
    st.markdown("**A. Exact claim as written**")
    for stage in analysis["counterfactual_source_echo"]["exact_claim_as_written"]:
        with st.container(border=True):
            st.markdown(f"**{stage['stage']}**")
            st.markdown(f"**Impact:** {stage['impact_label']}")
            st.markdown(f"**Source status badge:** {stage['source_status_badge']}")
            st.markdown(f"**Likely agent behavior:** {stage['likely_agent_behavior']}")
            st.markdown(f"**Why:** {stage['why']}")
            st.markdown(f"**Structural improvement needed:** {stage['structural_improvement_needed']}")
            st.markdown(f"**Source validation needed:** {stage['source_validation_needed']}")
    st.markdown("**B. Source-ready argument**")
    for stage in analysis["counterfactual_source_echo"]["recommended_upgraded_version"]:
        with st.container(border=True):
            st.markdown(f"**{stage['stage']}**")
            st.markdown(f"**Impact:** {stage['impact_label']}")
            st.markdown(f"**Source status badge:** {stage['source_status_badge']}")
            st.markdown(f"**Likely agent behavior:** {stage['likely_agent_behavior']}")
            st.markdown(f"**Why:** {stage['why']}")
            st.markdown(f"**Structural improvement needed:** {stage['structural_improvement_needed']}")
            st.markdown(f"**Source validation needed:** {stage['source_validation_needed']}")

with tabs[1]:
    list_cards(analysis["missing_evidence"], "piece", ["why_it_matters", "example_needed"])

with tabs[2]:
    rewrites = analysis["safer_claim_rewrites"]
    for title, value in [
        ("Safe marketing-language version", rewrites["safe_marketing_language"]),
        ("More substantiated version", rewrites["more_substantiated_version"]),
        (
            "Recommendation-ready version, assuming evidence exists",
            rewrites["recommendation_ready_version_assuming_evidence_exists"],
        ),
    ]:
        with st.container(border=True):
            st.markdown(f"**{title}**")
            st.code(value, language="markdown")

with tabs[3]:
    list_cards(
        analysis["triangulation_plan"],
        "surface",
        ["what_to_publish_or_collect", "why_an_ai_agent_would_care"],
    )

with tabs[4]:
    list_cards(
        analysis["red_team"],
        "risk",
        ["how_it_could_be_attacked_or_ignored", "how_to_neutralize"],
    )

with tabs[5]:
    evidence_yaml = evidence_block_as_yaml(analysis["agent_readable_evidence_block"])
    st.code(evidence_yaml, language="yaml")
    st.download_button(
        "Download evidence block",
        evidence_yaml,
        "agent-readable-evidence-block.yaml",
        "text/yaml",
    )

with tabs[6]:
    exports = export_markdown(analysis)
    for filename, content in exports.items():
        with st.container(border=True):
            st.markdown(f"**{filename}**")
            st.download_button(
                f"Download {filename}",
                content,
                filename,
                "text/plain",
                key=f"download-{filename}",
            )
