from __future__ import annotations

import asyncio
import json
import os
import uuid
from datetime import datetime, timezone

from dotenv import load_dotenv
from openai import AsyncOpenAI

from app.models import (
    DesignatedSitesSection,
    HabitatSection,
    PEAReport,
    SpeciesSection,
    SurveyRequest,
)

load_dotenv()

MODEL = "gpt-4o-mini"

SYSTEM_PROMPT = (
    "You are an expert ecological consultant writing a formal Preliminary "
    "Ecological Appraisal (PEA) for a UK ecological surveying company. Write "
    "in formal UK ecological report language. Be specific and reference the "
    "actual site data provided. Do not hallucinate species or designations "
    "not present in the input data. Follow standard UK ecological assessment "
    "methodology."
)

METHODOLOGY_INSTRUCTIONS = (
    "Reference EPSL/MAGIC database checks where appropriate. Use UKHab codes "
    "for habitat descriptions. Reference relevant UK legislation (Wildlife "
    "and Countryside Act 1981, NERC Act 2006, Conservation of Habitats "
    "Regulations 2017) where relevant. Include standard precautionary "
    "working method language where appropriate. Include biodiversity "
    "enhancement suggestions where relevant."
)

_client: AsyncOpenAI | None = None


def get_client() -> AsyncOpenAI:
    global _client
    if _client is None:
        api_key = os.environ.get("OPENAI_API_KEY")
        if not api_key:
            raise RuntimeError(
                "OPENAI_API_KEY is not set. Add it to the environment or a .env file."
            )
        _client = AsyncOpenAI(api_key=api_key)
    return _client


async def _call_structured(user_prompt: str, function_name: str, schema: dict) -> dict:
    client = get_client()
    response = await client.chat.completions.create(
        model=MODEL,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_prompt},
        ],
        tools=[
            {
                "type": "function",
                "function": {
                    "name": function_name,
                    "description": f"Record the {function_name} content for the PEA report.",
                    "parameters": schema,
                },
            }
        ],
        tool_choice={"type": "function", "function": {"name": function_name}},
    )
    tool_call = response.choices[0].message.tool_calls[0]
    return json.loads(tool_call.function.arguments)


HABITAT_SCHEMA = {
    "type": "object",
    "properties": {
        "site_context": {
            "type": "string",
            "description": "Brief description of the site's location and surrounding land use context.",
        },
        "ukhab_descriptions": {
            "type": "array",
            "items": {"type": "string"},
            "description": "On-site habitat descriptions, each prefixed with its UKHab code (e.g. 'g4 - Modified grassland: ...').",
        },
        "foreseen_impacts": {"type": "string"},
        "recommendations": {"type": "string"},
    },
    "required": ["site_context", "ukhab_descriptions", "foreseen_impacts", "recommendations"],
}

DESIGNATED_SITES_SCHEMA = {
    "type": "object",
    "properties": {
        "on_site_designations": {"type": "string"},
        "statutory_sites": {
            "type": "string",
            "description": "Statutory designated sites within 2km of the site.",
        },
        "non_statutory_sites": {
            "type": "string",
            "description": "Non-statutory designated sites within 2km of the site.",
        },
        "foreseen_impacts": {"type": "string"},
        "recommendations": {"type": "string"},
    },
    "required": [
        "on_site_designations",
        "statutory_sites",
        "non_statutory_sites",
        "foreseen_impacts",
        "recommendations",
    ],
}

SPECIES_SCHEMA = {
    "type": "object",
    "properties": {
        "epsl_data": {
            "type": "string",
            "description": "Relevant EPSL/MAGIC database search results, or a statement that none were found.",
        },
        "habitat_suitability": {"type": "string"},
        "survey_findings": {"type": "string"},
        "foreseen_impacts": {"type": "string"},
        "recommendations": {"type": "string"},
        "biodiversity_enhancements": {
            "type": "string",
            "description": "Suggested biodiversity enhancements, if relevant to this species group.",
        },
    },
    "required": [
        "epsl_data",
        "habitat_suitability",
        "survey_findings",
        "foreseen_impacts",
        "recommendations",
    ],
}


def _site_block(survey: SurveyRequest) -> str:
    site = survey.site_info
    return (
        f"Site name: {site.site_name}\n"
        f"Grid reference: {site.grid_reference}\n"
        f"Site area: {site.site_area_ha} ha\n"
        f"Development description: {site.development_description}\n"
    )


async def generate_habitat_section(survey: SurveyRequest) -> HabitatSection:
    prompt = (
        f"{_site_block(survey)}\n"
        f"Field notes on habitats and plants: {survey.habitat_notes}\n\n"
        "Write the Habitats and Plants section of the PEA. "
        f"{METHODOLOGY_INSTRUCTIONS}"
    )
    data = await _call_structured(prompt, "record_habitats_section", HABITAT_SCHEMA)
    return HabitatSection.model_validate(data)


async def generate_designated_sites_section(survey: SurveyRequest) -> DesignatedSitesSection:
    prompt = (
        f"{_site_block(survey)}\n"
        f"Field notes on locality and designated sites: {survey.designated_sites_notes}\n\n"
        "Write the Locality and Designated Sites section of the PEA. "
        f"{METHODOLOGY_INSTRUCTIONS}"
    )
    data = await _call_structured(
        prompt, "record_designated_sites_section", DESIGNATED_SITES_SCHEMA
    )
    return DesignatedSitesSection.model_validate(data)


async def generate_species_section(
    species_label: str, notes: str, survey: SurveyRequest, extra_instructions: str = ""
) -> SpeciesSection:
    prompt = (
        f"{_site_block(survey)}\n"
        f"Field notes on {species_label}: {notes}\n\n"
        f"Write the {species_label} section of the PEA. {extra_instructions} "
        f"{METHODOLOGY_INSTRUCTIONS}"
    )
    data = await _call_structured(
        prompt, "record_species_section", SPECIES_SCHEMA
    )
    return SpeciesSection.model_validate(data)


SPECIES_SPECS: list[tuple[str, str, str]] = [
    ("bats", "Bats", "Consider EPSL data, foraging and commuting habitat, and roosting habitat."),
    ("birds", "Birds", "Consider trees and vegetation, barn owls, and overwintering birds."),
    ("reptiles", "Reptiles", "Consider EPSL data and habitat suitability."),
    (
        "amphibians",
        "Amphibians",
        "Consider EPSL and survey data (including great crested newt) and habitat suitability.",
    ),
    ("badger", "Badger", ""),
    ("riparian", "Riparian Animals", ""),
    ("dormouse", "Hazel Dormouse", "Consider EPSL data and habitat suitability."),
    ("hedgehog", "Other (e.g. Hedgehog)", ""),
    ("invasive_species", "Invasive / Non-native species", ""),
    ("invertebrates", "Invertebrates", ""),
]


async def generate_pea_report(survey: SurveyRequest, survey_id: str | None = None) -> PEAReport:
    habitats_task = generate_habitat_section(survey)
    designated_sites_task = generate_designated_sites_section(survey)

    notes_by_field = {
        "bats": survey.bats_notes,
        "birds": survey.birds_notes,
        "reptiles": survey.reptiles_notes,
        "amphibians": survey.amphibians_notes,
        "badger": survey.badger_notes,
        "riparian": survey.riparian_notes,
        "dormouse": survey.dormouse_notes,
        "hedgehog": survey.hedgehog_notes,
        "invasive_species": survey.invasive_species_notes,
        "invertebrates": survey.invertebrates_notes,
    }

    species_tasks = {
        field: generate_species_section(label, notes_by_field[field], survey, extra)
        for field, label, extra in SPECIES_SPECS
    }

    habitats, designated_sites, *species_results = await asyncio.gather(
        habitats_task, designated_sites_task, *species_tasks.values()
    )
    species_by_field = dict(zip(species_tasks.keys(), species_results))

    return PEAReport(
        report_id=str(uuid.uuid4()),
        survey_id=survey_id or str(uuid.uuid4()),
        generated_at=datetime.now(timezone.utc).isoformat(),
        site_info=survey.site_info,
        author_info=survey.author_info,
        survey_conditions=survey.survey_conditions,
        habitats=habitats,
        designated_sites=designated_sites,
        invasive_species=species_by_field["invasive_species"],
        invertebrates=species_by_field["invertebrates"],
        bats=species_by_field["bats"],
        birds=species_by_field["birds"],
        reptiles=species_by_field["reptiles"],
        amphibians=species_by_field["amphibians"],
        badger=species_by_field["badger"],
        riparian=species_by_field["riparian"],
        dormouse=species_by_field["dormouse"],
        hedgehog=species_by_field["hedgehog"],
    )
