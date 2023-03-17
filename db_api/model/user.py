from .schedule_group_type_day import db



class User(db.Model):
    __tablename__ = 'users'
    id = db.Column(db.BigInteger, primary_key=True)
    first_name = db.Column(db.String(50))
    last_name = db.Column(db.String(50))
    username = db.Column(db.String(50))
    group = db.Column(db.String(15))
    daily_subscription_on = db.Column(db.Boolean())
    weekly_subscription_on = db.Column(db.Boolean())
    created_time = db.Column(DateTime(True), default=db.func.now())
    last_active_time = db.Column(DateTime(True), default=db.func.now())

