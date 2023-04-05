import pickle, os
from typing import Any
from module.settings import PATH_DUMP


def serializer(score: list, file_name: str) -> None:
    with open(os.path.join(PATH_DUMP, file_name), "wb") as fp:
        pickle.dump(score, fp)


def deserializer(file_name: str) -> Any:
    if not os.path.exists(os.path.join(PATH_DUMP, file_name)):
        return []
    with open(os.path.join(PATH_DUMP, file_name), "rb") as fp:
        b = pickle.load(fp)
        return b
