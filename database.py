import sqlite3
import os
import sys

def get_cwd():
    if getattr(sys, "frozen", False):
        path = os.path.dirname(sys.executable)
    elif __file__:
        path = os.path.dirname(__file__)
    else:
        path = os.path.dirname(os.getcwd())
    return path

def create_table():
    conn = sqlite3.connect(os.path.join(get_cwd(), 'mydatabase.db'))
    c = conn.cursor()
    sql = '''
            CREATE TABLE IF NOT EXISTS thm_form_data
            (
                nId                     INTEGER  PRIMARY KEY   AUTOINCREMENT,
                making_date             VARCHAR(20)     NOT NULL,
                client_name             VARCHAR(50)     NOT NULL,
                order_number            VARCHAR(20)     NOT NULL,
                production_number       VARCHAR(20)     NOT NULL,
                material_name           VARCHAR(20)     NOT NULL,
                quantity                INTEGER         NOT NULL,
                pro_no                  VARCHAR(4)      NOT NULL,
                z_heat                  VARCHAR(6)      NOT NULL,
                grind                   VARCHAR(4)      NOT NULL,
                chamfering              VARCHAR(4)      NOT NULL,
                cut_parameter           INTEGER         NOT NULL,
                machine_number          VARCHAR(10)     NOT NULL,
                delivery_date           VARCHAR(20)     NOT NULL,
                material_length         VARCHAR(20)     NOT NULL,
                material_width          VARCHAR(20)     NOT NULL,
                material_height         VARCHAR(20)     NOT NULL,
                finished_length         VARCHAR(20)     NOT NULL,
                finished_width          VARCHAR(20)     NOT NULL,
                finished_height         VARCHAR(20)     NOT NULL,
                fw                      VARCHAR(20)     NOT NULL,
                fl                      VARCHAR(20)     NOT NULL,
                processing_technology   VARCHAR(20)     NOT NULL,
                remark                  VARCHAR(200)    NOT NULL,
                pieces                  INTEGER         NOT NULL,
                weight                  VARCHAR(10)     NOT NULL,
                bale_check              INTEGER         NOT NULL,
                packet_check            INTEGER         NOT NULL,
                createdate              VARCHAR(20)     NOT NULL,
                updatedate              VARCHAR(20)     NOT NULL
            )
        '''
    c.execute(sql)

    sql = '''
            CREATE TABLE IF NOT EXISTS tvm_form_data
            (
                nId                     INTEGER  PRIMARY KEY   AUTOINCREMENT,
                making_date             VARCHAR(20)     NOT NULL,
                client_name             VARCHAR(50)     NOT NULL,
                order_number            VARCHAR(20)     NOT NULL,
                production_number       VARCHAR(20)     NOT NULL,
                material_name           VARCHAR(20)     NOT NULL,
                quantity                INTEGER         NOT NULL,
                pro_no                  VARCHAR(4)      NOT NULL,
                z_heat                  VARCHAR(6)      NOT NULL,
                grind                   VARCHAR(4)      NOT NULL,
                rou_magn                INTEGER         NOT NULL,
                fin_magn                INTEGER         NOT NULL,
                cut_parameter           INTEGER         NOT NULL,
                machine_number          VARCHAR(10)     NOT NULL,
                delivery_date           VARCHAR(20)     NOT NULL,
                material_length         VARCHAR(20)     NOT NULL,
                material_width          VARCHAR(20)     NOT NULL,
                material_height         VARCHAR(20)     NOT NULL,
                single_cut              VARCHAR(20)     NOT NULL,
                double_cut              VARCHAR(20)     NOT NULL,
                fh_single               VARCHAR(20)     NOT NULL,
                fh_double               VARCHAR(20)     NOT NULL,
                createdate              VARCHAR(20)     NOT NULL,
                updatedate              VARCHAR(20)     NOT NULL
            )
        '''
    c.execute(sql)

    conn.commit()
    conn.close()

def sql_build_dict(query, assoc_data = False):
    query.upper().strip()
    field = []
    value = []
    if type(assoc_data) == dict :
        return False

    if query == 'INSERT':
        for key , val in assoc_data.items():
            field.append(key)
            value.append(val)

        query = f" ( {','.join(field)} ) VALUES ( {','.join(value)} ) "
    elif query == 'UPDATE' :
        for key , val in assoc_data.items():
            value.append(f" {key} = {val} ")

        query = f" {','.join(value)} "

    return query

def insert_data(table,data):
    placeholders = ''
    values = ''
    conn = sqlite3.connect(os.path.join(get_cwd(), 'mydatabase.db'))
    c = conn.cursor()
    # 產生佔位符 確保新增時正確處理我的值
    placeholders = ', '.join('?' for _ in data)
    values = tuple(data.values())

    sql = f"INSERT INTO {table} ({', '.join(data.keys())}) VALUES ({placeholders})"
    c.execute(sql, values)

    conn.commit()
    conn.close()

def update_data(table, data, order_number):
    setholders = ''
    values = ''
    conn = sqlite3.connect(os.path.join(get_cwd(), 'mydatabase.db'))
    c = conn.cursor()

    set_clause = ', '.join(f"{key} = ?" for key in data)
    values = tuple(data.values())

    sql = f" UPDATE {table} SET {set_clause} WHERE order_number = '{order_number}' "

    c.execute(sql, values)

    conn.commit()
    conn.close()

def delete_data(table, id):
    conn = sqlite3.connect(os.path.join(get_cwd(), 'mydatabase.db'))
    c = conn.cursor()

    c.execute(f"DELETE FROM {table} WHERE nId=? LIMIT 1", (id,))

    conn.commit()
    conn.close()

def query_data(sql):
    conn = sqlite3.connect(os.path.join(get_cwd(), 'mydatabase.db'))
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    c.execute(sql)
    rows = c.fetchall()

    data = [row_to_dict(row) for row in rows]
    for dict in data:
        data = dict

    conn.close()

    return data

def row_to_dict(row):
    # 轉換成熟悉的dict型態
    return {key: row[key] for key in row.keys()}
