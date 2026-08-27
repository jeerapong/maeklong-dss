#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
แปลงผลลัพธ์จากแบบจำลอง WEAP ลุ่มน้ำแม่กลอง ให้เป็น data.json ที่หน้า Dashboard อ่านได้

นี่คือ *โครงร่าง* ที่ต้องเติมให้ตรงกับชื่อโหนดจริงในไฟล์ WEAP ของโครงการ
ทุกจุดที่ต้องแก้กำกับไว้ด้วย  # TODO

วิธีใช้
    python tools/weap_to_json.py --out data.json
    python tools/weap_to_json.py --out data.json --source weap     # อ่านจาก WEAP โดยตรง
    python tools/weap_to_json.py --out data.json --source api      # อ่านจาก REST API

--------------------------------------------------------------------------
ทางเลือกที่ 1: อ่านจาก WEAP โดยตรง (ต้องรันบน Windows ที่ติดตั้ง WEAP)
--------------------------------------------------------------------------
WEAP เปิด COM automation ให้เรียกจาก Python ได้ ผ่าน pywin32

    pip install pywin32
    import win32com.client
    WEAP = win32com.client.Dispatch("WEAP.WEAPApplication")
    WEAP.ActiveArea = "Mae Klong"
    WEAP.ActiveScenario = "Reference"
    val = WEAP.ResultValue(
        r'\\Demand Sites\\กำแพงแสน:Unmet Demand[Million Cubic Meter]',
        2569, 1, 2569, 12)      # ปีเริ่ม, เดือนเริ่ม, ปีจบ, เดือนจบ

เส้นทาง (branch path) ที่ใช้บ่อยในโครงการนี้
    \\Demand Sites\\<ชื่อ>:Water Demand[Million Cubic Meter]
    \\Demand Sites\\<ชื่อ>:Unmet Demand[Million Cubic Meter]
    \\Demand Sites\\<ชื่อ>:Coverage[Percent]
    \\Supply and Resources\\River\\<ชื่อ>\\Reservoirs\\<ชื่อ>:Storage Volume[Million Cubic Meter]
    \\Supply and Resources\\River\\<ชื่อ>\\Streamflow Gauges\\<ชื่อ>:Streamflow[Cubic Meter per Second]

--------------------------------------------------------------------------
ทางเลือกที่ 2: อ่านจาก REST API ของหน่วยงาน
--------------------------------------------------------------------------
ตั้งตัวแปรสภาพแวดล้อม SOURCE_URL และ SOURCE_TOKEN (อย่าเขียนค่าลงในไฟล์นี้)
"""

from __future__ import annotations

import argparse
import json
import os
from datetime import datetime, timezone, timedelta

TH_MONTHS = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.',
             'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.']

# สีของชุดข้อมูล — ต้องใช้ชื่อตัวแปร CSS เท่านั้น เพื่อให้สลับโหมดสว่าง/มืดได้ถูกต้อง
S1, S2, S3, S4, S5, S6 = (f'var(--s{i})' for i in range(1, 7))


def th_timestamp() -> str:
    """คืนเวลาปัจจุบันแบบไทย เช่น '27 ส.ค. 2569 06:00 น.'"""
    now = datetime.now(timezone(timedelta(hours=7)))
    return f'{now.day} {TH_MONTHS[now.month - 1]} {now.year + 543} {now:%H:%M} น.'


# ==========================================================================
# ทางเลือกที่ 1 — อ่านจาก WEAP
# ==========================================================================
def read_from_weap() -> dict:
    import win32com.client  # ต้องรันบน Windows ที่ติดตั้ง WEAP

    WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')
    WEAP.ActiveArea = 'Mae Klong'          # TODO ชื่อ Area จริงในไฟล์ WEAP
    year = 2569 - 543                       # WEAP ใช้ ค.ศ.

    def series(branch: str, unit: str) -> list[float]:
        """ดึงค่ารายเดือน 12 ค่าของ 1 ปี"""
        return [round(float(WEAP.ResultValue(f'{branch}[{unit}]', year, m, year, m)), 1)
                for m in range(1, 13)]

    # ---- ความต้องการน้ำแยกกิจกรรม -------------------------------------
    # TODO แก้ให้ตรงกับชื่อ Demand Site จริง และจัดกลุ่มตามกิจกรรม
    sector_map = [
        ('agri', 'เกษตรกรรม',    S1, [r'\Demand Sites\เกษตร แม่กลองใหญ่']),
        ('dom',  'อุปโภคบริโภค',  S2, [r'\Demand Sites\ประปา']),
        ('ind',  'อุตสาหกรรม',    S3, [r'\Demand Sites\อุตสาหกรรม']),
        ('com',  'พาณิชยกรรม',    S4, [r'\Demand Sites\พาณิชยกรรม']),
        ('sal',  'ผลักดันน้ำเค็ม', S5, [r'\Demand Sites\ผลักดันน้ำเค็ม']),
        ('eco',  'รักษาระบบนิเวศ', S6, [r'\Demand Sites\ระบบนิเวศ']),
    ]
    sectors = []
    for key, name, colour, branches in sector_map:
        total = [0.0] * 12
        for b in branches:
            for i, v in enumerate(series(f'{b}:Water Demand', 'Million Cubic Meter')):
                total[i] += v
        sectors.append({'k': key, 'n': name, 'c': colour,
                        'v': [round(v) for v in total]})

    # ---- น้ำขาดแคลนรวมทุก Demand Site ---------------------------------
    unmet = [0.0] * 12
    for _, _, _, branches in sector_map:
        for b in branches:
            for i, v in enumerate(series(f'{b}:Unmet Demand', 'Million Cubic Meter')):
                unmet[i] += v

    # ---- อ่างเก็บน้ำ ---------------------------------------------------
    # TODO แก้เส้นทางให้ตรงกับโครงข่ายจริง
    sri = series(r'\Supply and Resources\River\Khwae Yai\Reservoirs\Srinagarind'
                 r':Storage Volume', 'Million Cubic Meter')
    vjr = series(r'\Supply and Resources\River\Khwae Noi\Reservoirs\Vajiralongkorn'
                 r':Storage Volume', 'Million Cubic Meter')

    dead_sri, dead_vjr = 7763, 3292          # TODO ปริมาตรก้นอ่างจริง
    act = [round(a + b - dead_sri - dead_vjr) for a, b in zip(sri, vjr)]

    # เดือนที่ยังไม่ถึง ให้เป็น null เพื่อไม่ให้กราฟลากเส้นไปข้างหน้า
    this_month = datetime.now().month
    act = [v if i < this_month else None for i, v in enumerate(act)]

    return {
        'asOf': th_timestamp(),
        'RESERVOIRS': [
            {'n': 'เขื่อนศรีนครินทร์', 'cap': 9982,
             'cur': round(sri[this_month - 1] - dead_sri), 'rel': 14.2},   # TODO ค่าระบายจริง
            {'n': 'เขื่อนวชิราลงกรณ', 'cap': 5568,
             'cur': round(vjr[this_month - 1] - dead_vjr), 'rel': 11.6},
        ],
        'ACT': act,
        'SECTORS': sectors,
        'UNMET_BAU': [round(v) for v in unmet],
        # คีย์อื่น ๆ ไม่ต้องส่งก็ได้ — หน้าเว็บจะใช้ค่าเดิมไปก่อน
    }


# ==========================================================================
# ทางเลือกที่ 2 — อ่านจาก REST API
# ==========================================================================
def read_from_api() -> dict:
    import requests

    url = os.environ.get('SOURCE_URL')
    if not url:
        raise SystemExit('ยังไม่ได้ตั้งตัวแปรสภาพแวดล้อม SOURCE_URL')

    headers = {}
    token = os.environ.get('SOURCE_TOKEN')
    if token:
        headers['Authorization'] = f'Bearer {token}'

    res = requests.get(url, headers=headers, timeout=60)
    res.raise_for_status()
    payload = res.json()

    # TODO แปลงโครงสร้างของ API ให้ตรงกับคีย์ที่หน้า Dashboard ต้องการ (ดูตาราง ข้อ 5 ใน README)
    return {
        'asOf': th_timestamp(),
        'ACT': payload['reservoir_storage_monthly'],
        'UNMET_BAU': payload['unmet_demand_monthly'],
    }


# ==========================================================================
def main() -> None:
    ap = argparse.ArgumentParser(description='สร้าง data.json สำหรับ Dashboard แม่กลอง')
    ap.add_argument('--out', default='data.json', help='ไฟล์ปลายทาง')
    ap.add_argument('--source', choices=['weap', 'api'], default='api',
                    help='แหล่งข้อมูล: weap = อ่านจาก WEAP โดยตรง, api = อ่านจาก REST API')
    args = ap.parse_args()

    data = read_from_weap() if args.source == 'weap' else read_from_api()

    if not data.get('SECTORS') and not data.get('ACT'):
        raise SystemExit('ไม่พบข้อมูลที่ใช้ได้ — ยกเลิกการเขียนไฟล์ เพื่อไม่ให้ข้อมูลเดิมหาย')

    with open(args.out, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f'เขียน {args.out} แล้ว — {len(data)} คีย์ · ข้อมูล ณ {data["asOf"]}')


if __name__ == '__main__':
    main()
