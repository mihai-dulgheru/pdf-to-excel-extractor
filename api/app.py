from fastapi import FastAPI

from api.routes import router

app = FastAPI(title="PDF to Excel API", description="An API to process PDF invoices and generate Excel files.",
              version="1.0.0")

app.include_router(router)
