"""Conservative inline technical-token recognition; never rewrite source text.

This is not a shell parser. Unrecognized commands, Unicode/quoted paths, IPv6,
and ambiguous prose remain in the body font. Callers can disable recognition.
"""

import ipaddress
import re


_COMMANDS = frozenset(
    "ls pwd cd mkdir rmdir touch rm cp mv cat less head tail grep wc chmod chown "
    "chgrp umask useradd usermod userdel groupadd groupmod groupdel passwd sudo "
    "su bash sh vim vi nano find sed awk sort uniq cut tee tar gzip gunzip "
    "systemctl journalctl ip ping ssh scp rsync ps kill df du mount umount".split()
)
_COMMAND = re.compile(r"(?<![A-Za-z0-9_./-])[A-Za-z][A-Za-z0-9]*(?![A-Za-z0-9_./-])")
_CONTEXT = re.compile(r"命令|选项|参数|权限|终端|语法|执行|输入|板书|拆解|路径|文件|目录")
_OPTION = re.compile(r"(?<![A-Za-z0-9_-])--?[A-Za-z][A-Za-z0-9-]*(?![A-Za-z0-9_-])")
_PERMISSION = re.compile(
    r"(?<![A-Za-z0-9_-])(?:[bcdlps-])?(?:[r-][w-][xstST-]){3}(?![A-Za-z0-9_-])"
)
_PATH = re.compile(
    r"(?<![A-Za-z0-9_.:/~+-])(?:/|~/|\./|\.\./)"
    r"[A-Za-z0-9_.-]+(?:/[A-Za-z0-9_.-]+)*/?"
)
_IPV4 = re.compile(r"(?<![A-Za-z0-9_.])(?:\d{1,3}\.){3}\d{1,3}(?:/\d{1,2})?(?![A-Za-z0-9_./])")


def technical_spans(text):
    """Return non-overlapping ``(start, end)`` slices of ASCII technical tokens.

    Known command names need a technical context, command-style arguments, or
    to be the whole line. This avoids styling everyday English such as "cat".
    Options need a recognized command/technical context (or stand alone).
    """
    spans = []
    for pattern in (_PATH, _PERMISSION):
        for match in pattern.finditer(text):
            if pattern is _PERMISSION and not any(char != "-" for char in match.group()):
                continue  # A dashed prose separator is not a permission sample.
            # A prose full stop is not part of a path; keep it in the body run.
            end = match.end()
            if pattern is _PATH:
                end -= len(match.group()) - len(match.group().rstrip("."))
            if end > match.start():
                spans.append((match.start(), end))
    for match in _IPV4.finditer(text):
        try:
            ipaddress.IPv4Interface(match.group())
        except ValueError:
            continue
        spans.append(match.span())

    context = bool(_CONTEXT.search(text))
    commands = []
    for match in _COMMAND.finditer(text):
        if match.group() not in _COMMANDS:
            continue
        following = text[match.end():]
        command_args = bool(re.match(r"\s+(?:--?[A-Za-z]|[~./]|[0-7]{3,4}(?:\s|$))", following))
        if context or command_args or text.strip() == match.group():
            commands.append(match.span())
    spans.extend(commands)
    for match in _OPTION.finditer(text):
        if context or commands or text.strip() == match.group():
            spans.append(match.span())

    merged = []
    for start, end in sorted(spans):
        if merged and start <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(end, merged[-1][1]))
        else:
            merged.append((start, end))
    return merged


def split_technical_text(text):
    """Yield ``(verbatim_text, is_technical)`` chunks, preserving every character."""
    cursor = 0
    for start, end in technical_spans(text):
        if start > cursor:
            yield text[cursor:start], False
        yield text[start:end], True
        cursor = end
    if cursor < len(text) or not text:
        yield text[cursor:], False
