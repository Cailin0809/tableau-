#!/usr/bin/env python3
"""
typhoon_twbx_builder.py
=======================
一键生成 Tableau Public 可上传的台风路径 .twbx / .zip。

功能:
  1. 把任意 Excel (必须含 '台风信息' sheet) 读进 pandas
  2. 清洗移动速度 / 预报时间 / 现在时间 / 日期时间 等字段
  3. 用 Tableau Hyper API 生成 .hyper 数据提取 (database_version=1 以兼容旧版本)
  4. 生成符合 Tableau schema 的 .twb XML (worksheet: 台风路径, marks: Line, encoding: 台风编号/日期时间/经纬度/风速)
  5. 打包为 .twbx (Data/ + Data/Extracts/) 和 .zip (方便上传到雨课堂)

用法:
  python typhoon_twbx_builder.py \
      --excel /path/to/台风1949-2023.xlsx \
      --outdir /workspace/typhoon_out \
      --basename typhoon_path

依赖:
  pip install pandas openpyxl tableauhyperapi
"""

import argparse, os, sys, re, datetime, shutil, zipfile, tempfile
from pathlib import Path
import pandas as pd
from tableauhyperapi import (
    HyperProcess, Telemetry, Connection, CreateMode,
    TableDefinition, TableName, SqlType, Inserter,
)


# ---------- 1. 数据清洗 ----------
TEXT_COLS = ['移动方向', '台风级别', '国际名称', '中文名称']
INT_COLS  = ['现在气压', '预报时间', '现在时间', '现在风速', '台风编号']
REAL_COLS = ['移动速度', '现在纬度', '预报纬度', '预报气压', '预报经度', '现在经度']

def load_and_clean(excel_path: str) -> pd.DataFrame:
    df = pd.read_excel(excel_path, sheet_name='台风信息')
    for c in INT_COLS + REAL_COLS:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce')
    for c in INT_COLS:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce').astype('Int64')
    if '日期时间' in df.columns:
        df['日期时间'] = pd.to_datetime(df['日期时间'], errors='coerce')
    return df


# ---------- 2. .hyper 生成 ----------
def build_hyper(df: pd.DataFrame, hyper_path: str) -> None:
    os.makedirs(os.path.dirname(hyper_path), exist_ok=True)
    cols = [
        ('移动方向', SqlType.text()), ('台风级别', SqlType.text()),
        ('现在气压', SqlType.big_int()), ('移动速度', SqlType.double()),
        ('预报时间', SqlType.big_int()), ('国际名称', SqlType.text()),
        ('现在纬度', SqlType.double()), ('预报纬度', SqlType.double()),
        ('现在时间', SqlType.big_int()), ('中文名称', SqlType.text()),
        ('预报气压', SqlType.double()), ('预报经度', SqlType.double()),
        ('现在风速', SqlType.big_int()), ('现在经度', SqlType.double()),
        ('台风编号', SqlType.big_int()), ('日期时间', SqlType.timestamp()),
    ]
    text_cols = {n for n, t in cols if str(t).lower().startswith('text')}

    def conv(v):
        if v is None or (isinstance(v, float) and pd.isna(v)) or \
           (not isinstance(v, (str, int, float)) and pd.isna(v)):
            return None
        if isinstance(v, pd.Timestamp):
            return v.to_pydatetime()
        if hasattr(v, 'item'):
            try: return v.item()
            except Exception: return str(v)
        return v

    rows = []
    for _, r in df.iterrows():
        row = []
        for c, _ in cols:
            v = conv(r[c]) if c in r.index else None
            if c in text_cols and v is not None and not isinstance(v, str):
                v = str(v)
            row.append(v)
        rows.append(tuple(row))

    table_def = TableDefinition(
        table_name=TableName('Extract', 'Extract'),
        columns=[TableDefinition.Column(n, t) for n, t in cols],
    )
    if os.path.exists(hyper_path):
        os.remove(hyper_path)
    with HyperProcess(Telemetry.DO_NOT_SEND_USAGE_DATA_TO_TABLEAU,
                      parameters={'default_database_version': '1'}) as hyper:
        with Connection(hyper.endpoint, hyper_path, CreateMode.CREATE_AND_REPLACE) as conn:
            conn.catalog.create_schema('Extract')
            conn.catalog.create_table(table_def)
            with Inserter(conn, table_def) as ins:
                ins.add_rows(rows); ins.execute()
    print(f"[hyper] {hyper_path}  rows={len(rows)}  size={os.path.getsize(hyper_path):,}")


# ---------- 3. .twb 生成 ----------
TWB_TEMPLATE = """<?xml version='1.0' encoding='utf-8' ?>
<workbook original-version='18.1' source-build='2023.1.0 (20231.22.0703)'
          source-platform='win' version='18.1'
          xmlns:user='http://www.tableausoftware.com/xml/user'>
  <preferences>
    <preference name='ui.encoding.shelf.height' value='24' />
    <preference name='ui.shelf.height' value='26' />
  </preferences>
  <datasources>
    <datasource caption='{d_src}' inline='true' name='{d_src}' version='18.1'>
      <connection class='federated'>
        <named-connections>
          <named-connection caption='{d_src}' name='excel.{d_src}'>
            <connection class='excel-direct' filename='Data/{excel_file}'
                       password='' validate='yes' />
          </named-connection>
        </named-connections>
        <relation connection='excel.{d_src}' name='台风信息$' table='[台风信息$]' type='table' />
        <metadata-records>
{meta_block}
        </metadata-records>
      </connection>
      <aliases enabled='yes' />
{col_block}
      <extract count='-1' enabled='true' units='records'>
        <connection class='hyper' dbname='Data/Extracts/{hyper_file}'
                   schema='Extract' tablename='Extract' />
      </extract>
    </datasource>
  </datasources>
  <worksheets>
    <worksheet name='台风路径'>
      <table>
        <view>
          <datasource-dependencies datasource='{d_src}'>
            <column caption='台风级别' datatype='string' name='[台风级别]'
                    role='dimension' type='nominal' />
            <column-instance column='[台风级别]' derivation='None'
                             name='[none:台风级别:nk]' pivot='key' type='nominal' />
            <column caption='台风编号' datatype='integer' name='[台风编号]'
                    role='dimension' type='nominal' />
            <column-instance column='[台风编号]' derivation='None'
                             name='[none:台风编号:nk]' pivot='key' type='nominal' />
            <column caption='日期时间' datatype='datetime' name='[日期时间]'
                    role='dimension' type='quantitative' />
            <column-instance column='[日期时间]' derivation='None'
                             name='[none:日期时间:qk]' pivot='key' type='quantitative' />
            <column caption='现在风速' datatype='integer' name='[现在风速]'
                    role='measure' type='quantitative' />
            <column-instance column='[现在风速]' derivation='Sum'
                             name='[sum:现在风速:qk]' pivot='key' type='quantitative' />
            <column caption='现在经度' datatype='real' name='[现在经度]'
                    role='measure' type='quantitative' />
            <column-instance column='[现在经度]' derivation='Avg'
                             name='[avg:现在经度:qk]' pivot='key' type='quantitative' />
            <column caption='现在纬度' datatype='real' name='[现在纬度]'
                    role='measure' type='quantitative' />
            <column-instance column='[现在纬度]' derivation='Avg'
                             name='[avg:现在纬度:qk]' pivot='key' type='quantitative' />
          </datasource-dependencies>
          <aggregation value='true' />
        </view>
        <style>
          <style-rule element='mark'>
            <format attr='mark-labels-cull' value='true' />
          </style-rule>
        </style>
        <panes>
          <pane selection-relaxation-option='selection-relaxation-allow'>
            <mark class='Line' />
            <encodings>
              <color column='[{d_src}].[none:台风级别:nk]' />
              <size column='[{d_src}].[sum:现在风速:qk]' />
              <path column='[{d_src}].[none:日期时间:qk]' />
              <lod column='[{d_src}].[none:台风编号:nk]' />
            </encodings>
          </pane>
        </panes>
        <rows>[{d_src}].[avg:现在纬度:qk]</rows>
        <cols>[{d_src}].[avg:现在经度:qk]</cols>
      </table>
    </worksheet>
  </worksheets>
  <windows source-height='30'>
    <window class='worksheet' name='台风路径'>
      <cards>
        <edge name='left'>
          <strip size='160'>
            <card type='pages' />
            <card type='filters' />
            <card type='marks' />
          </strip>
        </edge>
        <edge name='top'>
          <strip size='2147483647'>
            <card type='columns' />
            <card type='rows' />
          </strip>
        </edge>
      </cards>
      <viewpoint><zoom type='entire-view' /></viewpoint>
    </window>
  </windows>
</workbook>
"""

FIELDS = [
    ('移动方向',   'string',  'Count',  0,  130,  'unicode'),
    ('台风级别',   'string',  'Count',  1,  130,  'unicode'),
    ('现在气压',   'integer', 'Sum',    2,    4,  None),
    ('移动速度',   'real',    'Sum',    3,    2,  None),
    ('预报时间',   'integer', 'Sum',    4,    4,  None),
    ('国际名称',   'string',  'Count',  5,  130,  'unicode'),
    ('现在纬度',   'real',    'Sum',    6,    2,  None),
    ('预报纬度',   'real',    'Sum',    7,    2,  None),
    ('现在时间',   'integer', 'Sum',    8,    4,  None),
    ('中文名称',   'string',  'Count',  9,  130,  'unicode'),
    ('预报气压',   'real',    'Sum',   10,    2,  None),
    ('预报经度',   'real',    'Sum',   11,    2,  None),
    ('现在风速',   'integer', 'Sum',   12,    4,  None),
    ('现在经度',   'real',    'Sum',   13,    2,  None),
    ('台风编号',   'integer', 'Count', 14,    4,  None),
    ('日期时间',   'datetime','Year',   15,    7,  None),
]

def _meta(fields):
    lines = []
    for name, _dt, agg, ord_, rt, td in fields:
        if td:
            lines.append(f"""          <metadata-record class='column'>
            <remote-name>{name}</remote-name>
            <remote-type>{rt}</remote-type>
            <local-name>[{name}]</local-name>
            <parent-name>[台风信息$]</parent-name>
            <remote-alias>{name}</remote-alias>
            <ordinal>{ord_}</ordinal>
            <local-type>{_dt}</local-type>
            <aggregation>{agg}</aggregation>
            <contains-null>true</contains-null>
            <attributes>
              <attribute datatype='string' name='text-domain' value='{td}' />
            </attributes>
          </metadata-record>""")
        else:
            lines.append(f"""          <metadata-record class='column'>
            <remote-name>{name}</remote-name>
            <remote-type>{rt}</remote-type>
            <local-name>[{name}]</local-name>
            <parent-name>[台风信息$]</parent-name>
            <remote-alias>{name}</remote-alias>
            <ordinal>{ord_}</ordinal>
            <local-type>{_dt}</local-type>
            <aggregation>{agg}</aggregation>
            <contains-null>true</contains-null>
          </metadata-record>""")
    return "\n".join(lines)

def _cols(fields):
    lines = []
    for name, dt, _agg, _o, _rt, _td in fields:
        role = 'dimension' if dt in ('string',) or (dt in ('integer',) and name in ('台风编号',)) else 'measure'
        typ  = 'nominal' if role == 'dimension' else 'quantitative'
        lines.append(f"      <column caption='{name}' datatype='{dt}' name='[{name}]' role='{role}' type='{typ}' />")
    return "\n".join(lines)

def build_twb(twb_path: str, d_src: str, excel_file: str, hyper_file: str) -> None:
    os.makedirs(os.path.dirname(twb_path), exist_ok=True)
    xml = TWB_TEMPLATE.format(
        d_src=d_src, excel_file=excel_file, hyper_file=hyper_file,
        meta_block=_meta(FIELDS), col_block=_cols(FIELDS),
    )
    # 合法性校验
    import xml.etree.ElementTree as ET
    ET.fromstring(xml.encode('utf-8'))
    Path(twb_path).write_text(xml, encoding='utf-8')
    print(f"[twb] validated XML  ->  {twb_path}")


# ---------- 4. .twbx / .zip 打包 ----------
def package(outdir: str, basename: str, excel_src: str, twb_path: str, hyper_path: str):
    data_dir = os.path.join(outdir, 'Data')
    ex_dir   = os.path.join(data_dir, 'Extracts')
    os.makedirs(ex_dir, exist_ok=True)
    excel_dst = os.path.join(data_dir, os.path.basename(excel_src))
    if not os.path.exists(excel_dst):
        shutil.copy2(excel_src, excel_dst)
    hyper_dst = os.path.join(ex_dir, os.path.basename(hyper_path))
    if os.path.abspath(hyper_path) != os.path.abspath(hyper_dst):
        shutil.copy2(hyper_path, hyper_dst)

    twbx = os.path.join(outdir, f'{basename}.twbx')
    z_path = os.path.join(outdir, f'{basename}.zip')
    rel_excel = os.path.join('Data', os.path.basename(excel_dst))
    rel_hyper = os.path.join('Data', 'Extracts', os.path.basename(hyper_dst))
    with zipfile.ZipFile(twbx, 'w', zipfile.ZIP_DEFLATED) as z:
        z.write(twb_path, f'{basename}.twb')
        z.write(excel_dst, rel_excel)
        z.write(hyper_dst, rel_hyper)
    with zipfile.ZipFile(z_path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.write(twbx, os.path.basename(twbx))
    print(f"[twbx] {twbx}  ({os.path.getsize(twbx):,} bytes)")
    print(f"[zip]  {z_path}  ({os.path.getsize(z_path):,} bytes)  ← 雨课堂可直接上传")


# ---------- main ----------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--excel',  required=True, help='源 Excel (含 台风信息 sheet)')
    ap.add_argument('--outdir', default='/workspace/typhoon_out')
    ap.add_argument('--basename', default='typhoon_path')
    args = ap.parse_args()

    os.makedirs(args.outdir, exist_ok=True)
    d_src   = 'federated.' + re.sub(r'[^A-Za-z0-9]', '', args.basename)[:12] or 'federated.t01'
    excel_file = 'typhoon.xlsx'
    hyper_file = 'typhoon.hyper'
    twb_path   = os.path.join(args.outdir, f'{args.basename}.twb')
    hyper_path = os.path.join(args.outdir, 'Data', 'Extracts', hyper_file)
    # 先 Data 目录建出来让 .twb 引用合法
    os.makedirs(os.path.join(args.outdir, 'Data'), exist_ok=True)
    dst_excel  = os.path.join(args.outdir, 'Data', excel_file)
    shutil.copy2(args.excel, dst_excel)

    print('\n=== 1) 数据清洗 ===')
    df = load_and_clean(dst_excel)
    print(f'    rows={len(df)}  台风={df["台风编号"].nunique()}  年份={df["日期时间"].dt.year.min()}-{df["日期时间"].dt.year.max()}')

    print('\n=== 2) Hyper 提取 ===')
    build_hyper(df, hyper_path)

    print('\n=== 3) .twb ===')
    build_twb(twb_path, d_src, excel_file, hyper_file)

    print('\n=== 4) 打包 .twbx + .zip ===')
    package(args.outdir, args.basename, dst_excel, twb_path, hyper_path)

    print(f'\n完成。输出目录: {args.outdir}')

if __name__ == '__main__':
    main()
