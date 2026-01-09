from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import List

import pandas as pd


@dataclass
class Business:
    name: str = ""
    address: str = ""
    website: str = ""
    phone_number: str = ""
    reviews_count: str = ""
    reviews_average: str = ""
    latitude: str = ""
    longitude: str = ""
    facebook: str = ""
    instagram: str = ""
    email: str = ""


@dataclass
class BusinessList:
    business_list: List[Business] = field(default_factory=list)
    save_at: str = "output"

    def dataframe(self) -> pd.DataFrame:
        columns = [
            "name",
            "address",
            "website",
            "phone_number",
            "reviews_count",
            "reviews_average",
            "latitude",
            "longitude",
            "facebook",
            "instagram",
            "email",
        ]
        data = [
            {
                "name": business.name,
                "address": business.address,
                "website": business.website,
                "phone_number": business.phone_number,
                "reviews_count": business.reviews_count,
                "reviews_average": business.reviews_average,
                "latitude": business.latitude,
                "longitude": business.longitude,
                "facebook": business.facebook,
                "instagram": business.instagram,
                "email": business.email,
            }
            for business in self.business_list
        ]
        return pd.DataFrame(data, columns=columns)

    def save_to_excel(self, filename: str) -> Path:
        output_dir = Path(self.save_at)
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / filename
        self.dataframe().to_excel(output_path, index=False)
        return output_path

    def save_to_csv(self, filename: str) -> Path:
        output_dir = Path(self.save_at)
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / filename
        self.dataframe().to_csv(output_path, index=False)
        return output_path
