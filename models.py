from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from datetime import datetime, date
from werkzeug.security import generate_password_hash, check_password_hash

db = SQLAlchemy()


class User(UserMixin, db.Model):
    __tablename__ = "users"
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(256), nullable=False)
    nume_complet = db.Column(db.String(200), nullable=True)
    is_admin = db.Column(db.Boolean, default=False)
    dark_mode = db.Column(db.Boolean, default=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)


class Firma(db.Model):
    __tablename__ = "firme"
    id = db.Column(db.Integer, primary_key=True)
    cod = db.Column(db.String(10), unique=True, nullable=False)
    nume = db.Column(db.String(200), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    angajati = db.relationship("ContractAngajat", back_populates="firma")

    def __repr__(self):
        return f"<Firma {self.cod} - {self.nume}>"


class Hotel(db.Model):
    __tablename__ = "hoteluri"
    id = db.Column(db.Integer, primary_key=True)
    nume = db.Column(db.String(200), unique=True, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    pontaje = db.relationship("Pontaj", back_populates="hotel")

    def __repr__(self):
        return f"<Hotel {self.nume}>"


class Angajat(db.Model):
    __tablename__ = "angajati"
    id = db.Column(db.Integer, primary_key=True)
    nume = db.Column(db.String(100), nullable=False)
    prenume = db.Column(db.String(100), nullable=False)
    nume_complet = db.Column(db.String(200), nullable=False, index=True)
    cnp = db.Column(db.String(13), unique=True, nullable=True)
    adresa = db.Column(db.String(500), nullable=True)
    telefon = db.Column(db.String(20), nullable=True)
    email = db.Column(db.String(200), nullable=True)
    activ = db.Column(db.Boolean, default=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    contracte = db.relationship(
        "ContractAngajat", back_populates="angajat", cascade="all, delete-orphan"
    )
    pontaje = db.relationship("Pontaj", back_populates="angajat")

    @property
    def firme_active(self):
        return [c.firma for c in self.contracte if c.activ]

    def __repr__(self):
        return f"<Angajat {self.nume_complet}>"


class ContractAngajat(db.Model):
    __tablename__ = "contracte"
    id = db.Column(db.Integer, primary_key=True)
    angajat_id = db.Column(db.Integer, db.ForeignKey("angajati.id"), nullable=False)
    firma_id = db.Column(db.Integer, db.ForeignKey("firme.id"), nullable=False)
    numar_contract = db.Column(db.String(50), nullable=True)
    functie = db.Column(db.String(200), nullable=True)
    cod_angajat = db.Column(db.String(50), nullable=True)
    tarif_orar = db.Column(db.Float, nullable=True)  # RON/ora
    data_inceput = db.Column(db.Date, nullable=True)
    data_sfarsit = db.Column(db.Date, nullable=True)
    activ = db.Column(db.Boolean, default=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    angajat = db.relationship("Angajat", back_populates="contracte")
    firma = db.relationship("Firma", back_populates="angajati")

    def genereaza_cod(self):
        if self.firma and self.numar_contract:
            self.cod_angajat = f"{self.firma.cod}-{self.numar_contract}"

    def __repr__(self):
        return f"<Contract {self.cod_angajat}>"


class Pontaj(db.Model):
    __tablename__ = "pontaje"
    id = db.Column(db.Integer, primary_key=True)
    angajat_id = db.Column(db.Integer, db.ForeignKey("angajati.id"), nullable=False)
    hotel_id = db.Column(db.Integer, db.ForeignKey("hoteluri.id"), nullable=False)
    data = db.Column(db.Date, nullable=False)
    ore = db.Column(db.Float, nullable=False)
    firma_cod = db.Column(db.String(10), nullable=True)
    saptamana = db.Column(db.String(20), nullable=True)
    fisier_sursa = db.Column(db.String(200), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    angajat = db.relationship("Angajat", back_populates="pontaje")
    hotel = db.relationship("Hotel", back_populates="pontaje")

    __table_args__ = (
        db.UniqueConstraint(
            "angajat_id", "data", "hotel_id", name="uq_pontaj_angajat_data_hotel"
        ),
    )

    def __repr__(self):
        return f"<Pontaj {self.angajat_id} {self.data} {self.ore}h>"


class AuditLog(db.Model):
    __tablename__ = "audit_log"
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True)
    actiune = db.Column(db.String(50), nullable=False)  # CREATE, UPDATE, DELETE, IMPORT, MERGE
    entitate = db.Column(db.String(50), nullable=False)  # Angajat, Pontaj, etc.
    entitate_id = db.Column(db.Integer, nullable=True)
    detalii = db.Column(db.Text, nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    user = db.relationship("User", backref="audit_logs")

    def __repr__(self):
        return f"<Audit {self.actiune} {self.entitate} #{self.entitate_id}>"


class Notification(db.Model):
    __tablename__ = "notifications"
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True)
    mesaj = db.Column(db.String(500), nullable=False)
    tip = db.Column(db.String(20), default="info")  # info, warning, success, danger
    citit = db.Column(db.Boolean, default=False)
    link = db.Column(db.String(200), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    user = db.relationship("User", backref="notifications")
