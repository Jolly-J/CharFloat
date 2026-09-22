#!/usr/bin/env python3
"""用例3：双客户端并发 + 目标锁隔离 + 审计串扰。"""
import json, sys, threading, time
sys.path.insert(0, '/Users/Python/Office Agent Bridge/.scratch/fault')
import mcp

WB = "钙钛矿各家企业现状.xlsx"
OUT = []


def log(*a):
    line = ' '.join(str(x) for x in a)
    OUT.append(line)
    print(line, flush=True)


def patch(sess, sheet, addr, values, tag):
    r = sess.call('excel_patch_cells', {
        'host': 'wps', 'workbookName': WB, 'sheetName': sheet,
        'address': addr, 'values': values, 'reason': tag})
    res = r.get('result', r)
    txt = res.get('content', [{}])[0].get('text', json.dumps(res, ensure_ascii=False))
    return res.get('isError'), txt


def main():
    A = mcp.Session('fault-A')
    B = mcp.Session('fault-B')
    log('A sessionId=%s  B sessionId=%s' % (A.sid, B.sid))

    # ---- 3.1 并发写不同工作表 ----
    results = {}

    def worker(sess, sheet, base, tag):
        for i in range(4):
            err, txt = patch(sess, sheet, 'A%d:B%d' % (base + i, base + i),
                             [[tag, i]], '%s-%d' % (tag, i))
            results[(tag, i)] = (err, txt)

    t0 = time.time()
    ta = threading.Thread(target=worker, args=(A, 'fault_case3_A', 1, 'A'))
    tb = threading.Thread(target=worker, args=(B, 'fault_case3_B', 1, 'B'))
    ta.start(); tb.start(); ta.join(); tb.join()
    log('并发用时 %.2fs，A 与 B 各 4 次写入完成' % (time.time() - t0))
    for k in sorted(results):
        err, txt = results[k]
        ok = '"success":true' in txt
        log('  %s -> isError=%s success=%s %s' % (k, err, ok, txt[:110]))

    # ---- 3.2 交叉读回：A 表内容是否被 B 污染 ----
    for sheet, tag in (('fault_case3_A', 'A'), ('fault_case3_B', 'B')):
        r = A.call('excel_read_range', {'host': 'wps', 'workbookName': WB,
                                        'sheetName': sheet, 'address': 'A1:B4'})
        txt = r['result']['content'][0]['text']
        vals = json.loads(txt)['values']
        tags = sorted({row[0] for row in vals if row and row[0]})
        log('交叉读回 %s 实际标签=%s (期望只有 %s)' % (sheet, tags, tag))

    # ---- 3.3 目标锁隔离 ----
    log('--- 锁隔离 ---')
    ra = A.call('wps_lock_target_document', {'component': 'excel', 'targetName': WB})
    log('A 锁 %s -> %s' % (WB, ra['result']['content'][0]['text'][:160]))
    rb = B.call('wps_lock_target_document', {'component': 'excel', 'targetName': 'B专属不存在.xlsx'})
    log('B 锁 B专属不存在.xlsx -> %s' % rb['result']['content'][0]['text'][:160])

    la = A.call('wps_get_locked_status', {})
    lb = B.call('wps_get_locked_status', {})
    log('A 看到的锁: %s' % la['result']['content'][0]['text'][:300])
    log('B 看到的锁: %s' % lb['result']['content'][0]['text'][:300])

    # 不传 workbookName，看各自解析到哪个文档
    na = A.call('excel_read_range', {'host': 'wps', 'sheetName': 'fault_case3_A', 'address': 'A1'})
    nb = B.call('excel_read_range', {'host': 'wps', 'sheetName': 'fault_case3_A', 'address': 'A1'})
    log('A 不传 workbookName 读 fault_case3_A: isError=%s %s' %
        (na['result'].get('isError'), na['result']['content'][0]['text'][:200]))
    log('B 不传 workbookName 读 fault_case3_A: isError=%s %s' %
        (nb['result'].get('isError'), nb['result']['content'][0]['text'][:200]))

    # ---- 3.4 审计串扰 ----
    log('--- 审计 ---')
    hist = A.call('wps_get_audit_history', {'workbookName': WB, 'limit': 30, 'view': 'summary'})
    recs = json.loads(hist['result']['content'][0]['text'])['records']
    for r in recs:
        if r['sheetName'] in ('fault_case3_A', 'fault_case3_B'):
            log('  audit id=%s sheet=%s clientName=%s desc=%s' %
                (r['id'][:8], r['sheetName'], r['clientName'], r['description']))

    # A 只筛自己的 sheet，看会不会带出 B 的记录
    ha = A.call('wps_get_audit_history', {'workbookName': WB, 'sheetName': 'fault_case3_A', 'limit': 30})
    ta_ = json.loads(ha['result']['content'][0]['text'])
    sheets = sorted({x['sheetName'] for x in ta_['records']})
    log('按 sheet=fault_case3_A 过滤后返回的表名集合: %s (total=%s)' % (sheets, ta_['total']))

    A.call('wps_unlock_target_document', {'component': 'excel'})
    B.call('wps_unlock_target_document', {'component': 'excel'})
    log('已解锁 A/B')


if __name__ == '__main__':
    main()
