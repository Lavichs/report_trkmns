import datetime
from datetime import timedelta
from pydantic_settings import BaseSettings, SettingsConfigDict


class Consts(BaseSettings):
    ECP_USER: str
    ECP_PASS: str

    ISLO_USER: str
    ISLO_PASS: str

    RABBIT_URL: str

    @property
    def TITLES(self) -> str:
        return [
            "Здание Детский стационар, ул. Комсомольская д. 200",
            "Здание Консультативно-диагностический центр, ул. Кобозева д. 25, к. А",
            "Здание Поликлиника №1, ул. Рыбаковская д.3, помещение 2",
            "Здание Поликлиника №2, ул. Пойменная д. 23, к. А",
            "Здание Поликлиника №3, ул. Алтайская д. 2",
            "Здание Поликлиника №4, ул. Туркестанская д. 43",
            "Здание Отделение медицинской реабилитации, ул. Кобозева д.12",
        ]

    @property
    def YESTERDAY(self) -> str:
        return (datetime.date.today() - timedelta(days=1)).strftime("%d.%m.%Y")

    @property
    def WEEK_AGO(self) -> str:
        return (datetime.date.today() - timedelta(days=7)).strftime("%d.%m.%Y")

    model_config = SettingsConfigDict(env_file=".env")


consts = Consts()
