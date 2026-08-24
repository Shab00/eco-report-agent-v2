from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from app.routes import health, report, survey

app = FastAPI(title="Greenlight Ecology PEA Report Generator")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(survey.router)
app.include_router(report.router)
app.include_router(health.router)
