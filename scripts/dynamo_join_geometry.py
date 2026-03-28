# -*- coding: utf-8 -*-
"""
THBIM - Join Geometry Manager for Dynamo
=========================================
IN[0]: Elements A (list)
IN[1]: Elements B (list)
IN[2]: Mode -> "join" | "switch" | "force_switch" | "unjoin" (default: "force_switch")

OUT[0]: Results   - ket qua tung cap
OUT[1]: Errors    - loi neu co
OUT[2]: Summary   - thong ke tong hop
"""

import clr

clr.AddReference('RevitServices')
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

doc = DocumentManager.Instance.CurrentDBDocument

# =========================
# INPUTS
# =========================
elementsA = IN[0]
elementsB = IN[1]
mode = (IN[2] or "force_switch").strip().lower()

VALID_MODES = ["join", "switch", "force_switch", "unjoin"]

# =========================
# HELPERS
# =========================
def to_list(x):
    if isinstance(x, list):
        return x
    return [x]

def unwrap_all(items):
    out = []
    for i in to_list(items):
        try:
            out.append(UnwrapElement(i))
        except:
            out.append(None)
    return out

def el_info(el):
    try:
        cat = el.Category.Name if el.Category else "NoCategory"
        return "ID:{} [{}]".format(el.Id.IntegerValue, cat)
    except:
        return "Invalid"

def can_use(el):
    return el is not None and hasattr(el, "Id") and el.Id != ElementId.InvalidElementId

# =========================
# VALIDATE MODE
# =========================
if mode not in VALID_MODES:
    OUT = [], ["Invalid mode '{}'. Use: {}".format(mode, ", ".join(VALID_MODES))], "ERROR"
else:
    # =========================
    # PREP
    # =========================
    listA = unwrap_all(elementsA)
    listB = unwrap_all(elementsB)
    count = min(len(listA), len(listB))

    results = []
    errors = []
    stats = {"success": 0, "skipped": 0, "failed": 0}

    TransactionManager.Instance.EnsureInTransaction(doc)

    for i in range(count):
        a = listA[i]
        b = listB[i]

        # --- Validate pair ---
        if not can_use(a) or not can_use(b):
            errors.append("Pair {}: Invalid element".format(i))
            stats["failed"] += 1
            continue

        if a.Id == b.Id:
            errors.append("Pair {}: Same element -> {}".format(i, el_info(a)))
            stats["skipped"] += 1
            continue

        try:
            joined = JoinGeometryUtils.AreElementsJoined(doc, a, b)
            pair = "{} <-> {}".format(el_info(a), el_info(b))

            if mode == "join":
                if not joined:
                    JoinGeometryUtils.JoinGeometry(doc, a, b)
                    results.append("Joined: " + pair)
                    stats["success"] += 1
                else:
                    results.append("Already joined: " + pair)
                    stats["skipped"] += 1

            elif mode == "switch":
                if joined:
                    JoinGeometryUtils.SwitchJoinOrder(doc, a, b)
                    results.append("Switched: " + pair)
                    stats["success"] += 1
                else:
                    results.append("Not joined, skip: " + pair)
                    stats["skipped"] += 1

            elif mode == "force_switch":
                if not joined:
                    JoinGeometryUtils.JoinGeometry(doc, a, b)
                JoinGeometryUtils.SwitchJoinOrder(doc, a, b)
                results.append("Force switched: " + pair)
                stats["success"] += 1

            elif mode == "unjoin":
                if joined:
                    JoinGeometryUtils.UnjoinGeometry(doc, a, b)
                    results.append("Unjoined: " + pair)
                    stats["success"] += 1
                else:
                    results.append("Already unjoined: " + pair)
                    stats["skipped"] += 1

        except Exception as e:
            errors.append("Pair {} FAILED: {} | {}".format(i, pair, str(e)))
            stats["failed"] += 1

    TransactionManager.Instance.TransactionTaskDone()

    # =========================
    # SUMMARY
    # =========================
    summary = "Mode: {} | Total: {} | Success: {} | Skipped: {} | Failed: {}".format(
        mode.upper(), count, stats["success"], stats["skipped"], stats["failed"]
    )

    OUT = results, errors, summary
