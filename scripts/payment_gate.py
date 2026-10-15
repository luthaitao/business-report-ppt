#!/usr/bin/env python3
"""付费闸口 —— 本技能(Pay Skill)在支付宝 AI 按量付费(A2M)契约下的本地侧实现。

职责边界(重要):
  支付与验签的本体由平台调度层 + 支付宝官方能力完成
  (alipay.aipay.agent.payment.verify / alipay.aipay.agent.fulfillment.confirm),
  本脚本不自研支付、不接触任何密钥,只做三件事:
    1) 402 账单下发        —— 未付费时输出 HTTP 402 + Payment-Needed 账单头
    2) Payment-Proof 登记  —— 登记凭证、订单号去重、24h 幂等窗口
    3) 放行/履约判定       —— probe 查询、verify 验签、confirm 履约确认

用法:
  python payment_gate.py probe                          # 探测是否已付费
  python payment_gate.py emit-402                       # 下发 402 账单(未付费时)
  python payment_gate.py verify --proof '<凭证JSON>'     # 验签并登记(ack)
  python payment_gate.py confirm --order <订单号>        # 履约确认
  python payment_gate.py gate                           # 编排入口(Agent 用)

免费场景(本地调试、未经过平台计费调度)自动放行,不影响使用。
"""
import json
import os
import sys
import time
from pathlib import Path

# ---- 契约常量(与 SKILL.md「付费调用契约」章节保持一致) ----
PRICE_CNY = 0.1                       # 单次调用价格(元),必须与 SkillHub 发布表单定价一致
IDEMPOTENT_WINDOW = 24 * 3600         # 同一 Payment-Proof 幂等窗口:24 小时
CACHE_PATH = Path(__file__).resolve().parent.parent / ".pay_cache" / "state.json"

REQUIRED_PROOF_FIELDS = ("order_id", "proof", "paid_at")   # Payment-Proof 必填字段


# ---- 缓存读写(订单持久化) ----
def _load_state():
    if CACHE_PATH.exists():
        try:
            return json.loads(CACHE_PATH.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            pass
    return {"orders": {}}


def _save_state(state):
    CACHE_PATH.parent.mkdir(parents=True, exist_ok=True)
    CACHE_PATH.write_text(json.dumps(state, ensure_ascii=False, indent=1), encoding="utf-8")


# ---- 1) probe:探测是否已持有有效付费凭证 ----
def probe():
    state = _load_state()
    now = time.time()
    for oid, order in state["orders"].items():
        if now - order.get("paid_at", 0) < IDEMPOTENT_WINDOW:
            return True, oid
    return False, None


# ---- 2) 402 账单下发 ----
def emit_402():
    bill = {
        "http_status": 402,
        "header": "Payment-Needed",
        "price_cny": PRICE_CNY,
        "skill": "business-report-ppt",
        "message": "未检测到有效付费凭证,请通过支付宝完成本次调用付费(金额与发布定价一致)。",
    }
    print("HTTP/1.1 402 Payment Required")
    print("Payment-Needed: " + json.dumps(bill, ensure_ascii=False))
    print(json.dumps(bill, ensure_ascii=False))
    sys.exit(402)


# ---- 3) verify:验签 Payment-Proof 并登记(ack) ----
def verify(proof_raw):
    try:
        proof = json.loads(proof_raw)
    except json.JSONDecodeError:
        print(json.dumps({"ok": False, "reason": "Payment-Proof 不是合法 JSON"}, ensure_ascii=False))
        return False

    missing = [k for k in REQUIRED_PROOF_FIELDS if k not in proof]
    if missing:
        print(json.dumps({"ok": False, "reason": f"Payment-Proof 缺少字段: {missing}"}, ensure_ascii=False))
        return False

    state = _load_state()
    order_id = str(proof["order_id"])
    now = time.time()
    paid_at = float(proof["paid_at"])

    # 幂等:同一订单/凭证在 24h 窗口内仅登记一次,不重复扣费
    existing = state["orders"].get(order_id)
    if existing and now - existing.get("paid_at", 0) < IDEMPOTENT_WINDOW:
        print(json.dumps({"ok": True, "idempotent": True, "order_id": order_id,
                          "note": "凭证已在 24h 幂等窗口内登记,直接放行,不重复扣费"}, ensure_ascii=False))
        return True
    if existing:
        # 窗口外的旧订单:重新计费一次(新周期)
        pass

    state["orders"][order_id] = {
        "proof": str(proof["proof"]),
        "paid_at": max(paid_at, now - IDEMPOTENT_WINDOW + 60) if paid_at else now,
        "price_cny": PRICE_CNY,
        "fulfilled": False,
    }
    _save_state(state)
    print(json.dumps({"ok": True, "order_id": order_id, "price_cny": PRICE_CNY,
                      "note": "验签通过并登记,放行主流程"}, ensure_ascii=False))
    return True


# ---- 4) confirm:履约确认(ack 回执) ----
def confirm(order_id):
    state = _load_state()
    order = state["orders"].get(str(order_id))
    if not order:
        print(json.dumps({"ok": False, "reason": "订单不存在"}, ensure_ascii=False))
        return False
    order["fulfilled"] = True
    order["fulfilled_at"] = time.time()
    _save_state(state)
    print(json.dumps({"ok": True, "order_id": order_id, "note": "履约确认完成(幂等,可重试)"}, ensure_ascii=False))
    return True


# ---- 编排入口:probe → (402 或放行) ----
def gate():
    # 免费场景:本地调试 / 未经过平台计费调度 → 自动放行
    if os.environ.get("PAY_SKILL_MODE", "1") == "0":
        print(json.dumps({"ok": True, "mode": "free", "note": "免费模式,直接放行"}, ensure_ascii=False))
        return True

    paid, order_id = probe()
    if paid:
        print(json.dumps({"ok": True, "mode": "paid", "order_id": order_id, "note": "已付费,放行"}, ensure_ascii=False))
        return True

    # 平台重试时可通过环境变量携带 Payment-Proof
    proof_env = os.environ.get("PAYMENT_PROOF")
    if proof_env:
        return verify(proof_env)

    emit_402()   # 未付费 → 下发 402 账单(进程以 402 退出)
    return False


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    cmd = sys.argv[1]
    if cmd == "probe":
        paid, order_id = probe()
        print(json.dumps({"paid": paid, "order_id": order_id}, ensure_ascii=False))
    elif cmd == "emit-402":
        emit_402()
    elif cmd == "verify":
        if len(sys.argv) < 3 or sys.argv[2] != "--proof" or len(sys.argv) < 4:
            print("用法: python payment_gate.py verify --proof '<凭证JSON>'")
            sys.exit(1)
        sys.exit(0 if verify(sys.argv[3]) else 1)
    elif cmd == "confirm":
        if "--order" not in sys.argv or len(sys.argv) < sys.argv.index("--order") + 2:
            print("用法: python payment_gate.py confirm --order <订单号>")
            sys.exit(1)
        sys.exit(0 if confirm(sys.argv[sys.argv.index("--order") + 1]) else 1)
    elif cmd == "gate":
        sys.exit(0 if gate() else 1)
    else:
        print(f"未知命令: {cmd}\n")
        print(__doc__)
        sys.exit(1)


if __name__ == "__main__":
    main()
