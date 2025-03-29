from dataclasses import dataclass


@dataclass(frozen=True)
class Constants:
    """
    Application constants for Excel formatting, PDF layout, and config paths.
    """
    __slots__ = ()

    APPLICATION_VERSION = "1.3.0"

    COLUMN_FORMATS = {"nr_crt": "General", "company": "@", "invoice_number": "0", "nc8_code": "@", "origin": "@",
                      "destination": "@", "invoice_value_eur": "#,##0.00", "net_weight": "#,##0",
                      "shipment_date": "dd.mmm", "exchange_rate": "#,##0.0000", "value_ron": "#,##0.00",
                      "vat_number": "@", "delivery_location": "0", "delivery_condition": "@", "percentage": "0.00",
                      "transport": "#,##0.00", "statistic": "#,##0.00", }

    COLUMNS = ["company", "invoice_number", "nc8_code", "origin", "destination", "invoice_value_eur", "net_weight",
               "shipment_date", "exchange_rate", "value_ron", "vat_number", "delivery_location", "delivery_condition"]

    FONT_SIZE = 12

    HEADERS = {"nr_crt": "Nr Crt", "company": "Firma", "invoice_number": "Nr Factura Marfa", "nc8_code": "Cod NC8",
               "origin": "Origine", "destination": "Destinatie", "invoice_value_eur": "Val Fact Euro",
               "net_weight": "Greutate Neta", "shipment_date": "Data Expeditiei", "exchange_rate": "Curs Valutar",
               "value_ron": "Valoare Ron", "vat_number": "Vat Cumparator", "delivery_location": "Loc Livrare",
               "delivery_condition": "Conditie Livrare", "percentage": "%", "transport": "Transport",
               "statistic": "Statistica", }

    """
    +-------------------------------+
    |           section_1           |
    |                               |
    +-------------------------------+
    |   section_2   |   section_3   |
    |               |               |
    |               |               |
    |               |               |
    |               |               |
    +-------------------------------+
    |           section_4           |
    |                               |
    |                               |
    |                               |
    |                               |
    |                               |
    +-------------------------------+
    |           section_5           |
    +-------------------------------+
    """
    PROPORTIONS = {"section_1": (0.0, 0.0, 1.0, 0.16), "section_2": (0.0, 0.16, 0.46, 0.54),
                   "section_3": (0.46, 0.16, 1.0, 0.54), "section_4": (0.0, 0.54, 1.0, 0.93),
                   "section_5": (0.0, 0.93, 1.0, 1.0), }

    SCALING_FACTOR = 1.2
    CONFIG_FILE = "config.json"
    DEFAULT_CODE = 2093

    LOCATION_MAPPING = {"BUDESTI": 1759, "CATEASCA": 1826, "CRAIOVA": 1593, }
