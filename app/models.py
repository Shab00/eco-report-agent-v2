from __future__ import annotations

from pydantic import BaseModel


class SurveyConditions(BaseModel):
    date: str
    temperature_c: float
    humidity_percent: float
    cloud_cover_percent: float
    wind_kmh: float
    rain: str


class SiteInfo(BaseModel):
    site_name: str
    grid_reference: str
    site_area_ha: float
    development_description: str
    prepared_for: str
    doc_ref: str


class AuthorInfo(BaseModel):
    name: str
    credentials: str
    role: str


class SpeciesSection(BaseModel):
    epsl_data: str
    habitat_suitability: str
    survey_findings: str
    foreseen_impacts: str
    recommendations: str
    biodiversity_enhancements: str | None = None


class HabitatSection(BaseModel):
    site_context: str
    ukhab_descriptions: list[str]
    foreseen_impacts: str
    recommendations: str


class DesignatedSitesSection(BaseModel):
    on_site_designations: str
    statutory_sites: str
    non_statutory_sites: str
    foreseen_impacts: str
    recommendations: str


class PEAReport(BaseModel):
    report_id: str
    survey_id: str
    generated_at: str
    site_info: SiteInfo
    author_info: AuthorInfo
    survey_conditions: SurveyConditions
    habitats: HabitatSection
    designated_sites: DesignatedSitesSection
    invasive_species: SpeciesSection
    invertebrates: SpeciesSection
    bats: SpeciesSection
    birds: SpeciesSection
    reptiles: SpeciesSection
    amphibians: SpeciesSection
    badger: SpeciesSection
    riparian: SpeciesSection
    dormouse: SpeciesSection
    hedgehog: SpeciesSection
    docx_path: str | None = None


class SurveyRequest(BaseModel):
    """Raw survey submission: structured site/author/condition data plus
    free-text ecologist field notes per topic. The generator expands the
    notes into the full PEAReport section models via GPT-4o-mini."""

    site_info: SiteInfo
    author_info: AuthorInfo
    survey_conditions: SurveyConditions
    habitat_notes: str
    designated_sites_notes: str
    invasive_species_notes: str
    invertebrates_notes: str
    bats_notes: str
    birds_notes: str
    reptiles_notes: str
    amphibians_notes: str
    badger_notes: str
    riparian_notes: str
    dormouse_notes: str
    hedgehog_notes: str


class SurveyResponse(BaseModel):
    report_id: str
    generated_at: str
    download_url: str
