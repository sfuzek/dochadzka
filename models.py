from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from datetime import datetime

db = SQLAlchemy()

class User(UserMixin, db.Model):
    __tablename__ = 'users'
    id         = db.Column(db.Integer, primary_key=True)
    meno       = db.Column(db.String(100), nullable=False)
    email      = db.Column(db.String(150), unique=True, nullable=False)
    heslo      = db.Column(db.String(255), nullable=False)
    is_admin   = db.Column(db.Boolean, default=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    klienti  = db.relationship('Klient', backref='user', lazy=True, cascade='all, delete-orphan')
    zaznamy  = db.relationship('Zaznam', backref='user', lazy=True, cascade='all, delete-orphan')
    odoslane_ulohy  = db.relationship('Uloha', foreign_keys='Uloha.odosielatel_id', backref='odosielatel', lazy=True)
    prijate_ulohy   = db.relationship('Uloha', foreign_keys='Uloha.prijimatel_id',  backref='prijimatel',  lazy=True)


class Klient(db.Model):
    __tablename__ = 'klienti'
    id         = db.Column(db.Integer, primary_key=True)
    nazov      = db.Column(db.String(200), nullable=False)
    user_id    = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    aktivny    = db.Column(db.Boolean, default=True)

    zaznamy = db.relationship('Zaznam', backref='klient', lazy=True)


class Zaznam(db.Model):
    __tablename__ = 'zaznamy'
    id        = db.Column(db.Integer, primary_key=True)
    user_id   = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    klient_id = db.Column(db.Integer, db.ForeignKey('klienti.id'), nullable=False)
    datum     = db.Column(db.Date, nullable=False)
    cas_start = db.Column(db.DateTime, nullable=False)
    cas_stop  = db.Column(db.DateTime, nullable=True)
    poznamka  = db.Column(db.Text, default='')

    cinnosti = db.relationship('Cinnost', backref='zaznam', lazy=True,
                               cascade='all, delete-orphan', order_by='Cinnost.cas_od')

    @property
    def trvanie_sekundy(self):
        if self.cas_stop:
            return int((self.cas_stop - self.cas_start).total_seconds())
        return None

    @property
    def trvanie_format(self):
        sek = self.trvanie_sekundy
        if sek is None:
            return 'Prebieha...'
        return f'{sek // 3600:02d}:{(sek % 3600) // 60:02d}'


class Cinnost(db.Model):
    __tablename__ = 'cinnosti'
    id        = db.Column(db.Integer, primary_key=True)
    zaznam_id = db.Column(db.Integer, db.ForeignKey('zaznamy.id'), nullable=False)
    popis     = db.Column(db.Text, nullable=False)
    cas_od    = db.Column(db.DateTime, nullable=False)
    cas_do    = db.Column(db.DateTime, nullable=True)

    @property
    def trvanie_format(self):
        if self.cas_do:
            sek = int((self.cas_do - self.cas_od).total_seconds())
            return f'{sek // 3600:02d}:{(sek % 3600) // 60:02d}'
        return 'Prebieha...'


class Uloha(db.Model):
    __tablename__ = 'ulohy'
    id              = db.Column(db.Integer, primary_key=True)
    odosielatel_id  = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    prijimatel_id   = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    text            = db.Column(db.Text, nullable=False)
    urgencia        = db.Column(db.Integer, default=2)  # 1=nízka 2=stredná 3=vysoká
    hotovo          = db.Column(db.Boolean, default=False)
    hotovo_at       = db.Column(db.DateTime, nullable=True)
    created_at      = db.Column(db.DateTime, default=datetime.utcnow)
