#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
install_skill: 一键将 teaching-plan Skill 安装或更新到全局或指定 Agent 技能目录。
"""

import sys
import os
import shutil
import argparse

STANDARD_GLOBAL_SKILLS_DIR = os.path.expanduser("~/.agents/skills")
SOURCE_SKILL_DIR = os.path.join(os.path.dirname(__file__), "skills", "teaching-plan")


def detect_default_skills_dir():
    """
    智能探测本机已存在的 Agent 全局技能目录，默认使用通用 ~/.agents/skills
    """
    candidates = [
        os.path.expanduser("~/.agents/skills"),
        os.path.expanduser("~/.claude/skills"),
        os.path.expanduser("~/.gemini/config/skills"),
    ]
    for path in candidates:
        if os.path.isdir(path):
            return path
    return STANDARD_GLOBAL_SKILLS_DIR


def install_skill(dest_root=None, local=False):
    if not os.path.exists(SOURCE_SKILL_DIR):
        print(f"❌ 错误: 未找到源 Skill 目录: {SOURCE_SKILL_DIR}", file=sys.stderr)
        sys.exit(1)

    if dest_root:
        target_root = os.path.expanduser(dest_root)
    elif local:
        target_root = os.path.abspath(".agents/skills")
    else:
        target_root = detect_default_skills_dir()

    target_skill_dir = os.path.join(target_root, "teaching-plan")

    print(f"📦 准备安装 teaching-plan Skill...")
    print(f"   源目录:   {SOURCE_SKILL_DIR}")
    print(f"   目标目录: {target_skill_dir}")

    os.makedirs(target_root, exist_ok=True)

    if os.path.exists(target_skill_dir):
        print(f"   ℹ️ 目标目录已存在，正在覆盖更新...")
        shutil.rmtree(target_skill_dir)

    shutil.copytree(SOURCE_SKILL_DIR, target_skill_dir)
    print(f"✨ 安装成功！teaching-plan Skill 已部署至: {target_skill_dir}")
    print(f"🎉 现在您可以在任意项目中通过自然语言调用该 Skill 生成教案！")


def main():
    parser = argparse.ArgumentParser(description="teaching-plan Skill 一键安装工具")
    parser.add_argument(
        "--dest",
        default=None,
        help=f"指定目标技能根目录 (默认自动探测或使用: {STANDARD_GLOBAL_SKILLS_DIR})"
    )
    parser.add_argument(
        "--local",
        action="store_true",
        help="安装到当前项目局部目录 (.agents/skills/teaching-plan)"
    )
    args = parser.parse_args()
    install_skill(dest_root=args.dest, local=args.local)


if __name__ == "__main__":
    main()
