# !/usr/bin/env python
# coding:utf-8

import os
import json


class FetchRealLevel():
    def __init__(self):
        self.allow_db_list = ["oncokb", "civic", "cgi"]
        self.func_list = ["self.check_oncokb", "self.check_civic", "self.check_cgi"]

    def level(self, db_name, level, origin_descr, base_descr):
        if db_name not in self.allow_db_list:
            return False
        if isinstance(base_descr, str):
            base_descr = base_descr.decode("utf-8")
        if isinstance(origin_descr, str):
            origin_descr = origin_descr.decode("utf-8")
        origin_descr = origin_descr.strip()
        base_descr = base_descr.strip()
        origin_descr_list = self.name_transform(db_name, origin_descr)
        if origin_descr_list is False:
            return False
        trans_args = [level, origin_descr_list, base_descr]
        real_level = eval(self.func_list[self.allow_db_list.index(db_name)])(*trans_args)
        return real_level

    def name_transform(self, db_name, origin_descr):
        flag, disease_info = self.load_json("static/base_data/disease_name.json")
        origin_descr_list = origin_descr.split(",") if "," in origin_descr else [origin_descr]
        target_list = list()
        for index_n, index_value in enumerate(origin_descr_list):
            index_value = index_value.strip()
            if index_value == "CANCER":
                continue
            if index_value not in disease_info[db_name].keys():
                return False
            target_list.extend(disease_info[db_name][index_value])
        target_list = map(unicode, target_list)
        return flag if flag is False else target_list

    def check_oncokb(self, level, origin_descr_list, base_descr):
        if level not in ["1", "2A", "3A", "4", "R1"]:
            return False
        elif level in ["1", "2A", "R1"]:
            return "A" if base_descr in origin_descr_list else "C"
        elif level == "3A":
            return "B" if base_descr in origin_descr_list else "C"
        else:
            return "D"

    def check_civic(self, level, origin_descr_list, base_descr):
        if level not in ["A: Validated", "B: Clinical evidence", "C: Case study",
                         "D: Preclinical evidence", "E: Indirect evidence"]:
            return False
        elif level.startswith("A"):
            return "A" if base_descr in origin_descr_list else "C"
        elif level.startswith("B"):
            return "C"
        else:
            return "D"

    def check_cgi(self, level, origin_descr_list, base_descr):
        if level not in ["FDA guidelines", "NCCN guidelines", "NCCN/CAP guidelines", "Late trials",
                         "Early trials", "Case report", "Pre-clinical"]:
            return False
        elif level in ["FDA guidelines", "NCCN guidelines", "NCCN/CAP guidelines"]:
            return "A" if base_descr in origin_descr_list else "C"
        elif level == "Late trials":
            return "B" if base_descr in origin_descr_list else "C"
        elif level == "Early trials":
            return "C"
        else:
            return "D"

    def load_json(self, json_file):
        if os.path.isfile(json_file) is False:
            return False, "json file not exist"
        read_json = open(json_file)
        json_content = read_json.read()
        read_json.close()
        try:
            json_desc = json.loads(json_content, "utf-8")
        except ValueError as ve:
            return False, "File content not json"
        return True, json_desc




