# database.py
"""
数据库操作类
负责所有数据的增删改查、统计、备份等功能
"""

import sqlite3
import os
import shutil
from datetime import datetime

# 数据库文件存放在程序同目录下
DB_PATH = os.path.join(os.path.dirname(__file__), 'equipment_db.sqlite')


class Database:
    def __init__(self):
        """连接数据库，启用外键约束，初始化所有表"""
        self.conn = sqlite3.connect(DB_PATH)
        self.conn.row_factory = sqlite3.Row  # 让查询结果可以通过列名访问
        self.cursor = self.conn.cursor()
        # 启用外键约束（防止删除器材后留下孤立记录）
        self.cursor.execute("PRAGMA foreign_keys = ON")
        self.init_tables()

        #数据清洗
        self.cursor.execute("UPDATE equipment SET quality_level='堪用' WHERE quality_level='堪用品'")
        self.cursor.execute("UPDATE equipment SET quality_level='待修' WHERE quality_level='待修品'")
        self.cursor.execute("UPDATE equipment SET quality_level='报废' WHERE quality_level='报废品'")
        self.conn.commit()

    def init_tables(self):
        """创建所有需要的数据库表（如果不存在的话）"""
        # 器材档案表
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS equipment (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                serial_no TEXT,          -- 器材编号
                name TEXT NOT NULL,      -- 器材名称（必填）
                spec_model TEXT NOT NULL,-- 规格型号（必填）
                category_code TEXT,      -- 品种标识码
                unit TEXT,               -- 计量单位
                quality_level TEXT DEFAULT '新品',  -- 质量等级
                manufacturer TEXT,       -- 生产厂家
                unit_price REAL DEFAULT 0, -- 单价
                production_date TEXT,    -- 生产日期
                storage_life INTEGER,    -- 存储寿命（月）
                source_type TEXT,        -- 来源类别
                pricing_method TEXT,     -- 计价方法
                contract_no TEXT,        -- 合同编号
                belong_unit TEXT,        -- 所属单位
                belong_equipment TEXT,   -- 所属装备
                created_at TEXT DEFAULT (datetime('now','localtime')), -- 创建时间
                updated_at TEXT DEFAULT (datetime('now','localtime'))  -- 更新时间
            )
        ''')

        # 机关入库记录表
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS org_stock_in (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                equipment_id INTEGER NOT NULL,  -- 对应器材ID
                quantity REAL NOT NULL DEFAULT 1, -- 数量
                total_price REAL DEFAULT 0,      -- 总价
                in_time TEXT,                    -- 入库时间
                stock_in_person TEXT,            -- 入库人
                operator TEXT,                   -- 经办人
                remark TEXT,                     -- 备注
                batch_no TEXT,                   -- 批次号
                FOREIGN KEY (equipment_id) REFERENCES equipment(id)
            )
        ''')

        # 机关出库记录表
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS org_stock_out (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                equipment_id INTEGER NOT NULL,
                to_unit TEXT NOT NULL,         -- 发放给哪个中队
                out_direction TEXT,            -- 出库去向
                quantity REAL NOT NULL DEFAULT 1,
                total_price REAL DEFAULT 0,
                out_time TEXT,
                stock_out_person TEXT,
                operator TEXT,
                remark TEXT,
                batch_no TEXT,
                FOREIGN KEY (equipment_id) REFERENCES equipment(id)
            )
        ''')

        # 中队出入库记录表（入库和出库都记录在这里，用 record_type 区分）
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS unit_stock_record (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                equipment_id INTEGER NOT NULL,
                belong_unit TEXT NOT NULL,     -- 所属中队
                record_type TEXT NOT NULL DEFAULT 'in',  -- 'in'=入库, 'out'=出库
                quantity REAL NOT NULL DEFAULT 1,
                total_price REAL DEFAULT 0,
                in_time TEXT,                  -- 入库时间（入库时填写）
                out_time TEXT,                 -- 出库时间（出库时填写）
                out_direction TEXT,            -- 出库去向
                source TEXT,                   -- 来源（如"机关下拨"）
                stock_in_person TEXT,
                stock_out_person TEXT,
                operator TEXT,
                remark TEXT,
                batch_no TEXT,
                FOREIGN KEY (equipment_id) REFERENCES equipment(id)
            )
        ''')

        # 操作日志表
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS operation_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                action TEXT NOT NULL,          -- 操作类型（新增/编辑/删除/入库/出库等）
                target_type TEXT,              -- 操作对象类型
                target_id INTEGER,             -- 操作对象ID
                detail TEXT,                   -- 详细信息（包含器材名称、数量等）
                operator TEXT DEFAULT '',      -- 操作人
                role TEXT DEFAULT '',          -- 角色
                belong_unit TEXT DEFAULT '',   -- 所属单位
                created_at TEXT DEFAULT (datetime('now','localtime'))  -- 操作时间
            )
        ''')

        # 创建索引，提高查询速度
        self.cursor.execute("CREATE INDEX IF NOT EXISTS idx_equip_name_spec ON equipment(name, spec_model)")
        self.cursor.execute("CREATE INDEX IF NOT EXISTS idx_org_in_equip ON org_stock_in(equipment_id)")
        self.cursor.execute("CREATE INDEX IF NOT EXISTS idx_org_out_equip ON org_stock_out(equipment_id)")
        self.cursor.execute("CREATE INDEX IF NOT EXISTS idx_unit_record ON unit_stock_record(equipment_id, belong_unit, record_type)")

        self.conn.commit()

    # ==================== 事务管理 ====================
    def begin_transaction(self):
        """开始事务（批量操作前调用）"""
        self.cursor.execute("BEGIN")

    def commit_transaction(self):
        """提交事务（所有操作成功时调用）"""
        self.conn.commit()

    def rollback_transaction(self):
        """回滚事务（操作失败时调用，撤销所有未提交的更改）"""
        self.conn.rollback()

    # ==================== 器材档案操作 ====================
    def add_equipment(self, data: dict):
        """新增器材，返回新器材的ID"""
        sql = '''INSERT INTO equipment 
            (serial_no, name, spec_model, category_code, unit, quality_level,
             manufacturer, unit_price, production_date, storage_life, source_type,
             pricing_method, contract_no, belong_unit, belong_equipment)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)'''
        self.cursor.execute(sql, (
            data.get('serial_no', ''), data['name'], data['spec_model'],
            data.get('category_code', ''), data.get('unit', ''),
            data.get('quality_level', '新品'), data.get('manufacturer', ''),
            data.get('unit_price', 0), data.get('production_date', ''),
            data.get('storage_life', None), data.get('source_type', ''),
            data.get('pricing_method', ''), data.get('contract_no', ''),
            data.get('belong_unit', ''), data.get('belong_equipment', '')
        ))
        self.conn.commit()
        return self.cursor.lastrowid

    def update_equipment(self, equip_id: int, data: dict):
        """更新器材信息"""
        fields = []
        values = []
        # 允许更新的字段列表
        allowed = ['serial_no', 'name', 'spec_model', 'category_code', 'unit',
                   'quality_level', 'manufacturer', 'unit_price', 'production_date',
                   'storage_life', 'source_type', 'pricing_method', 'contract_no',
                   'belong_unit', 'belong_equipment']
        for k in allowed:
            if k in data:
                fields.append(f"{k}=?")
                values.append(data[k])
        if not fields:
            return
        # 自动更新修改时间
        fields.append("updated_at=?")
        values.append(datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
        values.append(equip_id)
        sql = f"UPDATE equipment SET {','.join(fields)} WHERE id=?"
        self.cursor.execute(sql, values)
        self.conn.commit()

    def delete_equipment(self, equip_id: int):
        """删除器材（因为有外键约束，有库存记录的器材无法删除）"""
        self.cursor.execute("DELETE FROM equipment WHERE id=?", (equip_id,))
        self.conn.commit()

    def search_equipment(self, keyword='', belong_unit='', quality_level='', limit=500):
        """按关键词、单位、质量等级搜索器材（质量等级改为模糊匹配）"""
        sql = "SELECT * FROM equipment WHERE 1=1"
        params = []
        if keyword:
            sql += " AND (belong_unit LIKE ? OR name LIKE ? OR spec_model LIKE ? OR serial_no LIKE ?)"
            kw = f"%{keyword}%"
            params.extend([kw, kw, kw, kw])
        if belong_unit:
            sql += " AND belong_unit LIKE ?"
            params.append(f"%{belong_unit}%")
        if quality_level:
            # 改为模糊匹配，输入“堪用”可匹配“堪用品”、“堪用”
            sql += " AND quality_level LIKE ?"
            params.append(f"%{quality_level}%")
        sql += " ORDER BY updated_at DESC LIMIT ?"
        params.append(limit)
        self.cursor.execute(sql, params)
        return self.cursor.fetchall()

    def get_equipment_by_id(self, equip_id: int):
        """根据ID获取器材信息"""
        self.cursor.execute("SELECT * FROM equipment WHERE id=?", (equip_id,))
        return self.cursor.fetchone()

    # ==================== 库存计算 ====================
    def get_org_current_stock(self, equipment_id: int):
        """计算某器材在机关的当前库存（入库总量 - 出库总量）"""
        self.cursor.execute("SELECT COALESCE(SUM(quantity),0) FROM org_stock_in WHERE equipment_id=?", (equipment_id,))
        in_total = self.cursor.fetchone()[0]
        self.cursor.execute("SELECT COALESCE(SUM(quantity),0) FROM org_stock_out WHERE equipment_id=?", (equipment_id,))
        out_total = self.cursor.fetchone()[0]
        return in_total - out_total

    def get_unit_current_stock(self, equipment_id: int, belong_unit: str):
        """计算某器材在某中队的当前库存（入库总量 - 出库总量）"""
        self.cursor.execute(
            "SELECT COALESCE(SUM(quantity),0) FROM unit_stock_record WHERE equipment_id=? AND belong_unit=? AND record_type='in'",
            (equipment_id, belong_unit))
        in_total = self.cursor.fetchone()[0]
        self.cursor.execute(
            "SELECT COALESCE(SUM(quantity),0) FROM unit_stock_record WHERE equipment_id=? AND belong_unit=? AND record_type='out'",
            (equipment_id, belong_unit))
        out_total = self.cursor.fetchone()[0]
        return in_total - out_total

    # ==================== 机关入库 ====================
    def add_org_stock_in(self, data: dict):
        """添加入库记录"""
        sql = '''INSERT INTO org_stock_in 
            (equipment_id, quantity, total_price, in_time, stock_in_person, operator, remark, batch_no)
            VALUES (?,?,?,?,?,?,?,?)'''
        self.cursor.execute(sql, (
            data['equipment_id'], data['quantity'], data['total_price'],
            data['in_time'], data.get('stock_in_person', ''),
            data.get('operator', ''), data.get('remark', ''),
            data.get('batch_no', '')
        ))
        # 注意：这里不再自动提交，由业务层控制事务
        return self.cursor.lastrowid

    def get_all_org_stock_in(self):
        """获取所有机关入库记录（联合器材表）"""
        sql = '''SELECT si.*, e.serial_no, e.name, e.spec_model, e.unit, e.unit_price,
                        e.manufacturer, e.contract_no, e.belong_unit, e.belong_equipment,
                        e.quality_level, e.category_code, e.production_date, e.storage_life,
                        e.source_type, e.pricing_method
                 FROM org_stock_in si LEFT JOIN equipment e ON si.equipment_id = e.id
                 ORDER BY si.in_time DESC'''
        self.cursor.execute(sql)
        return self.cursor.fetchall()

    def get_org_stock_in_by_date(self, in_date: str):
        """获取指定日期的机关入库记录"""
        sql = '''SELECT si.*, e.serial_no, e.name, e.spec_model, e.unit, e.unit_price,
                        e.manufacturer, e.contract_no, e.belong_unit, e.belong_equipment,
                        e.quality_level, e.category_code, e.production_date, e.storage_life,
                        e.source_type, e.pricing_method
                 FROM org_stock_in si LEFT JOIN equipment e ON si.equipment_id = e.id
                 WHERE si.in_time = ? ORDER BY si.id'''
        self.cursor.execute(sql, (in_date,))
        return self.cursor.fetchall()

    # ==================== 机关出库 ====================
    def add_org_stock_out(self, data: dict):
        """添加机关出库记录"""
        sql = '''INSERT INTO org_stock_out 
            (equipment_id, to_unit, out_direction, quantity, total_price, out_time, stock_out_person, operator, remark, batch_no)
            VALUES (?,?,?,?,?,?,?,?,?,?)'''
        self.cursor.execute(sql, (
            data['equipment_id'], data['to_unit'], data.get('out_direction', ''),
            data['quantity'], data['total_price'], data['out_time'],
            data.get('stock_out_person', ''), data.get('operator', ''),
            data.get('remark', ''), data.get('batch_no', '')
        ))
        return self.cursor.lastrowid

    def get_all_org_stock_out(self):
        """获取所有机关出库记录"""
        sql = '''SELECT so.*, e.serial_no, e.name, e.spec_model, e.unit, e.unit_price,
                        e.manufacturer, e.contract_no, e.belong_unit, e.belong_equipment,
                        e.quality_level, e.category_code, e.production_date, e.storage_life,
                        e.source_type, e.pricing_method
                 FROM org_stock_out so LEFT JOIN equipment e ON so.equipment_id = e.id
                 ORDER BY so.out_time DESC'''
        self.cursor.execute(sql)
        return self.cursor.fetchall()

    def get_org_stock_out_by_date(self, out_date: str):
        """获取指定日期的机关出库记录"""
        sql = '''SELECT so.*, e.serial_no, e.name, e.spec_model, e.unit, e.unit_price,
                        e.manufacturer, e.contract_no, e.belong_unit, e.belong_equipment,
                        e.quality_level, e.category_code, e.production_date, e.storage_life,
                        e.source_type, e.pricing_method
                 FROM org_stock_out so LEFT JOIN equipment e ON so.equipment_id = e.id
                 WHERE so.out_time = ? ORDER BY so.id'''
        self.cursor.execute(sql, (out_date,))
        return self.cursor.fetchall()

    # ==================== 中队记录 ====================
    def add_unit_record(self, data: dict):
        """添加中队出入库记录（入库和出库都用这个函数，通过 record_type 区分）"""
        sql = '''INSERT INTO unit_stock_record 
            (equipment_id, belong_unit, record_type, quantity, total_price, 
             in_time, out_time, out_direction, source, stock_in_person, stock_out_person, operator, remark, batch_no)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)'''
        self.cursor.execute(sql, (
            data['equipment_id'], data['belong_unit'], data['record_type'],
            data['quantity'], data['total_price'],
            data.get('in_time', ''), data.get('out_time', ''),
            data.get('out_direction', ''), data.get('source', ''),
            data.get('stock_in_person', ''), data.get('stock_out_person', ''),
            data.get('operator', ''), data.get('remark', ''),
            data.get('batch_no', '')
        ))
        return self.cursor.lastrowid

    def get_unit_records(self, belong_unit: str, record_type: str = ''):
        """获取中队出入库记录"""
        if record_type:
            sql = '''SELECT ur.*, e.serial_no, e.name, e.spec_model, e.unit, e.unit_price,
                            e.manufacturer, e.contract_no, e.belong_unit as equip_belong_unit, e.belong_equipment,
                            e.quality_level, e.category_code, e.production_date, e.storage_life,
                            e.source_type, e.pricing_method
                     FROM unit_stock_record ur LEFT JOIN equipment e ON ur.equipment_id = e.id
                     WHERE ur.belong_unit=? AND ur.record_type=? 
                     ORDER BY ur.in_time DESC, ur.out_time DESC'''
            self.cursor.execute(sql, (belong_unit, record_type))
        else:
            sql = '''SELECT ur.*, e.serial_no, e.name, e.spec_model, e.unit, e.unit_price,
                            e.manufacturer, e.contract_no, e.belong_unit as equip_belong_unit, e.belong_equipment,
                            e.quality_level, e.category_code, e.production_date, e.storage_life,
                            e.source_type, e.pricing_method
                     FROM unit_stock_record ur LEFT JOIN equipment e ON ur.equipment_id = e.id
                     WHERE ur.belong_unit=?
                     ORDER BY ur.in_time DESC, ur.out_time DESC'''
            self.cursor.execute(sql, (belong_unit,))
        return self.cursor.fetchall()

    # ==================== 库存统计 ====================
    def get_org_stock_summary(self):
        """机关库存总览统计（器材种类、入库总量、出库总量、当前库存，以及按质量等级分布）"""
        self.cursor.execute('''
            SELECT e.quality_level, COUNT(e.id) as count,
                   COALESCE(SUM(si.quantity),0) - COALESCE(SUM(so.quantity),0) as total_stock
            FROM equipment e
            LEFT JOIN (SELECT equipment_id, SUM(quantity) as quantity FROM org_stock_in GROUP BY equipment_id) si 
                ON e.id = si.equipment_id
            LEFT JOIN (SELECT equipment_id, SUM(quantity) as quantity FROM org_stock_out GROUP BY equipment_id) so 
                ON e.id = so.equipment_id
            GROUP BY e.quality_level
        ''')
        quality_stats = self.cursor.fetchall()
        self.cursor.execute('SELECT COUNT(*) FROM equipment')
        total_types = self.cursor.fetchone()[0]
        self.cursor.execute('SELECT COALESCE(SUM(quantity),0) FROM org_stock_in')
        total_in = self.cursor.fetchone()[0]
        self.cursor.execute('SELECT COALESCE(SUM(quantity),0) FROM org_stock_out')
        total_out = self.cursor.fetchone()[0]
        return {'quality_stats': quality_stats, 'total_types': total_types,
                'total_in': total_in, 'total_out': total_out, 'current_stock': total_in - total_out}

    def get_unit_stock_summary(self, belong_unit: str):
        """中队库存总览统计"""
        self.cursor.execute('''
            SELECT e.quality_level, COUNT(e.id) as count,
                   COALESCE(SUM(CASE WHEN ur.record_type='in' THEN ur.quantity ELSE 0 END),0) 
                   - COALESCE(SUM(CASE WHEN ur.record_type='out' THEN ur.quantity ELSE 0 END),0) as total_stock
            FROM equipment e
            LEFT JOIN unit_stock_record ur ON e.id = ur.equipment_id AND ur.belong_unit = ?
            GROUP BY e.quality_level
        ''', (belong_unit,))
        quality_stats = self.cursor.fetchall()
        self.cursor.execute('SELECT COUNT(*) FROM equipment')
        total_types = self.cursor.fetchone()[0]
        self.cursor.execute("SELECT COALESCE(SUM(quantity),0) FROM unit_stock_record WHERE belong_unit=? AND record_type='in'", (belong_unit,))
        total_in = self.cursor.fetchone()[0]
        self.cursor.execute("SELECT COALESCE(SUM(quantity),0) FROM unit_stock_record WHERE belong_unit=? AND record_type='out'", (belong_unit,))
        total_out = self.cursor.fetchone()[0]
        return {'quality_stats': quality_stats, 'total_types': total_types,
                'total_in': total_in, 'total_out': total_out, 'current_stock': total_in - total_out}

    def get_org_low_stock(self, threshold=5):
        """获取机关库存低于阈值的器材列表"""
        self.cursor.execute('''
            SELECT e.*, COALESCE(si.total_in,0) - COALESCE(so.total_out,0) as current_stock
            FROM equipment e
            LEFT JOIN (SELECT equipment_id, SUM(quantity) as total_in FROM org_stock_in GROUP BY equipment_id) si 
                ON e.id = si.equipment_id
            LEFT JOIN (SELECT equipment_id, SUM(quantity) as total_out FROM org_stock_out GROUP BY equipment_id) so 
                ON e.id = so.equipment_id
            WHERE current_stock <= ? AND current_stock > 0 ORDER BY current_stock ASC
        ''', (threshold,))
        return self.cursor.fetchall()

    def get_unit_low_stock(self, belong_unit: str, threshold=5):
        """获取中队库存低于阈值的器材列表"""
        self.cursor.execute('''
            SELECT e.*,
                   COALESCE(SUM(CASE WHEN ur.record_type='in' THEN ur.quantity ELSE 0 END),0) 
                   - COALESCE(SUM(CASE WHEN ur.record_type='out' THEN ur.quantity ELSE 0 END),0) as current_stock
            FROM equipment e
            LEFT JOIN unit_stock_record ur ON e.id = ur.equipment_id AND ur.belong_unit = ?
            GROUP BY e.id HAVING current_stock <= ? AND current_stock > 0 ORDER BY current_stock ASC
        ''', (belong_unit, threshold))
        return self.cursor.fetchall()

    # ==================== 到期提醒 ====================
    def get_expiring_soon(self, days=30):
        """获取存储寿命将在指定天数内到期的器材"""
        sql = '''SELECT e.*, date(e.production_date, '+' || e.storage_life || ' months') as expire_date
                 FROM equipment e
                 WHERE e.storage_life IS NOT NULL AND e.storage_life > 0 
                   AND e.production_date IS NOT NULL AND e.production_date != ''
                   AND date(e.production_date, '+' || e.storage_life || ' months') <= date('now', '+' || ? || ' days')
                   AND date(e.production_date, '+' || e.storage_life || ' months') >= date('now')'''
        self.cursor.execute(sql, (days,))
        return self.cursor.fetchall()

    # ==================== 操作日志 ====================
    def add_log(self, action: str, target_type: str = '', target_id: int = None,
                detail: str = '', operator: str = '', role: str = '', belong_unit: str = ''):
        """记录操作日志"""
        self.cursor.execute(
            "INSERT INTO operation_log (action,target_type,target_id,detail,operator,role,belong_unit) VALUES (?,?,?,?,?,?,?)",
            (action, target_type, target_id, detail, operator, role, belong_unit))
        self.conn.commit()  # 日志即时提交

    def get_logs(self, action: str = '', role: str = '', belong_unit: str = '',
                 date_from: str = '', date_to: str = '', limit: int = 500):
        """查询操作日志"""
        sql = "SELECT * FROM operation_log WHERE 1=1"
        params = []
        if action: sql += " AND action=?"; params.append(action)
        if role: sql += " AND role=?"; params.append(role)
        if belong_unit: sql += " AND belong_unit LIKE ?"; params.append(f"%{belong_unit}%")
        if date_from: sql += " AND date(created_at) >= ?"; params.append(date_from)
        if date_to: sql += " AND date(created_at) <= ?"; params.append(date_to)
        sql += " ORDER BY created_at DESC LIMIT ?"; params.append(limit)
        self.cursor.execute(sql, params)
        return self.cursor.fetchall()

    def clear_logs_before(self, before_date: str):
        """清理指定日期之前的日志"""
        self.cursor.execute("DELETE FROM operation_log WHERE date(created_at) < ?", (before_date,))
        deleted = self.cursor.rowcount
        self.conn.commit()
        return deleted

    # ==================== 备份恢复 ====================
    def backup_database(self, backup_path: str) -> bool:
        """在线备份数据库（不中断当前连接）"""
        # 使用第二个连接进行备份
        backup_conn = sqlite3.connect(backup_path)
        self.conn.backup(backup_conn)
        backup_conn.close()
        return True

    def restore_database(self, backup_path: str) -> bool:
        """恢复数据库（覆盖当前数据库文件，需要重启程序）"""
        self.conn.close()
        shutil.copy2(backup_path, DB_PATH)
        # 重新连接
        self.conn = sqlite3.connect(DB_PATH)
        self.conn.row_factory = sqlite3.Row
        self.cursor = self.conn.cursor()
        self.cursor.execute("PRAGMA foreign_keys = ON")
        return True

    def get_database_info(self):
        """获取数据库文件信息"""
        if os.path.exists(DB_PATH):
            size = os.path.getsize(DB_PATH)
            mtime = os.path.getmtime(DB_PATH)
            return {
                'path': DB_PATH,
                'size': size,
                'size_mb': round(size/1024/1024, 2),
                'modified': datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M:%S')
            }
        return None

    def close(self):
        """关闭数据库连接"""
        self.conn.close()