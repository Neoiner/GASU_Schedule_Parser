from typing import List
import sqlalchemy as sa
from gino import Gino

db = Gino()

class ScheduleGroupTypeDay(db.Model):
    __tablename__ = 'lessons_new'
    group_name = db.Column(db.Text(30), primary_key=True)
    discipline = db.Column(db.Text(255))
    format = db.Column(db.Text(10))
    number_of_less = db.Column(db.SmallInteger())
    week = db.Column(db.Boolean(), primary_key=True)
    day_of_week = db.Column(db.SmallInteger(), primary_key=True)
    professor = db.Column(db.Text())
    classroom = db.Column(db.Text(30))

    _pk = db.PrimaryKeyConstraint('group_name', 'week', 'day_of_week', name='lessons_pkey')

    def __str__(self):
        model = self.__class__.__name__
        table: sa.Table = sa.inspect(self.__class__)
        primary_key_columns: List[sa.Column] = table.columns
        values = {
            column.name: getattr(self, self._column_name_map[column.name])
            for column in primary_key_columns
        }
        values_str = " ".join(f"{name}={value!r}" for name, value in values.items())
        return f"<{model} {values_str}>"