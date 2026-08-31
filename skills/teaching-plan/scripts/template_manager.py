# -*- coding: utf-8 -*-
"""
TemplateManager: Loads, manages, and resolves declarative document templates.
Provides lightweight template expression evaluation and template registry.
"""

import os
import json
import re
import copy

def _resolve_default_template_path():
    # 1. 尝试当前目录下的 templates
    p1 = os.path.join(os.path.dirname(__file__), "templates", "default_vocational.json")
    if os.path.exists(p1):
        return p1
    # 2. 尝试上级目录下的 templates
    p2 = os.path.join(os.path.dirname(__file__), "..", "templates", "default_vocational.json")
    if os.path.exists(p2):
        return os.path.abspath(p2)
    return p1

DEFAULT_TEMPLATE_PATH = _resolve_default_template_path()

class TemplateManager:
    def __init__(self, template_config=None):
        """
        Initialize TemplateManager with template configuration.
        :param template_config: Path to template JSON, template name, or template dict.
        """
        if template_config is None:
            self.template = self._load_template_file(DEFAULT_TEMPLATE_PATH)
        elif isinstance(template_config, dict):
            self.template = copy.deepcopy(template_config)
        elif isinstance(template_config, str):
            # Check if it's a direct file path
            if os.path.exists(template_config):
                self.template = self._load_template_file(template_config)
            else:
                # Check if it's a name in templates/ directory (both local and parent)
                local_named = os.path.join(os.path.dirname(__file__), "templates", f"{template_config}.json")
                parent_named = os.path.join(os.path.dirname(__file__), "..", "templates", f"{template_config}.json")
                if os.path.exists(local_named):
                    self.template = self._load_template_file(local_named)
                elif os.path.exists(parent_named):
                    self.template = self._load_template_file(parent_named)
                else:
                    raise FileNotFoundError(f"未找到模板文件: {template_config} (尝试路径: {parent_named})")
        else:
            raise ValueError(f"不支持的模板配置类型: {type(template_config)}")

    def _load_template_file(self, filepath):
        with open(filepath, "r", encoding="utf-8") as f:
            return json.load(f)

    def get_section(self, *keys, default=None):
        """Access nested template definition."""
        curr = self.template
        for k in keys:
            if isinstance(curr, dict) and k in curr:
                curr = curr[k]
            else:
                return default
        return curr

    @staticmethod
    def evaluate_expression(expr_str, context, max_depth=5):
        """
        Evaluates simple template expressions like:
        - "{{ lesson.title }}"
        - "{{ item.stage_index }}\n{{ item.stage_name }}"
        - "{{ cover.title | default('教  学  设  计') }}"
        - Raw text strings without template tags
        Supports iterative evaluation of nested tags.
        """
        if not isinstance(expr_str, str):
            return expr_str

        curr_str = expr_str
        pattern = re.compile(r"\{\{\s*(.*?)\s*\}\}")

        for _ in range(max_depth):
            if "{{" not in curr_str:
                break

            def replace_match(match):
                token = match.group(1).strip()
                
                # Check for default filter: e.g. cover.title | default('...')
                default_val = ""
                if "|" in token:
                    parts = token.split("|", 1)
                    var_path = parts[0].strip()
                    filter_part = parts[1].strip()
                    def_match = re.search(r"default\((?:'|\")(.*?)(?:'|\")\)", filter_part)
                    if def_match:
                        default_val = def_match.group(1)
                else:
                    var_path = token

                # Traverse context
                keys = var_path.split(".")
                curr = context
                for k in keys:
                    if isinstance(curr, dict) and k in curr:
                        curr = curr[k]
                    else:
                        curr = None
                        break

                if curr is None or curr == "":
                    return default_val
                return str(curr)

            # If the whole string is just a single tag and resolves to an object/list
            stripped = curr_str.strip()
            single_tag_match = re.fullmatch(r"\{\{\s*(.*?)\s*\}\}", stripped)
            if single_tag_match:
                raw_token = single_tag_match.group(1).strip()
                if "|" not in raw_token:
                    keys = raw_token.split(".")
                    curr = context
                    found = True
                    for k in keys:
                        if isinstance(curr, dict) and k in curr:
                            curr = curr[k]
                        else:
                            found = False
                            break
                    if found and curr is not None and not isinstance(curr, str):
                        return curr

            prev_str = curr_str
            curr_str = pattern.sub(replace_match, curr_str)
            if prev_str == curr_str:
                break

        return curr_str
