from flask import Flask, render_template, request, redirect, url_for, flash, send_file, jsonify, session
from flask_login import LoginManager, login_user, logout_user, login_required, current_user
from flask_wtf.csrf import CSRFProtect, CSRFError
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address
from models import db, User, Klient, Zaznam, Cinnost, Uloha, Nastavenie, LoginLog, Dovolenka, PredefinovanaCinnost, ZiadostOZmenu
from datetime import datetime, timezone, timedelta
from collections import defaultdict
import bcrypt
import io
import os
import json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'dochadzka-dev-key-zmenit')
database_url = os.environ.get('DATABASE_URL', 'sqlite:///dochadzka.db')
if 'postgresql' in database_url or database_url.startswith('postgres://'):
    database_url = database_url.replace('postgres://', 'postgresql+pg8000://')
    database_url = database_url.replace('postgresql://', 'postgresql+pg8000://')
    if 'pg8000' not in database_url:
        database_url = database_url.replace('postgresql://', 'postgresql+pg8000://')
app.config['SQLALCHEMY_DATABASE_URI'] = database_url
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['WTF_CSRF_TIME_LIMIT'] = 3600  # CSRF token platí 1 hodinu

db.init_app(app)

# CSRF ochrana
csrf = CSRFProtect(app)

# Rate limiter — ukladá pokusy v pamäti (pre prod možno redis, ale memory stačí)
limiter = Limiter(
    key_func=get_remote_address,
    app=app,
    default_limits=[],
    storage_uri='memory://'
)

login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'
login_manager.login_message = 'Pre prístup sa prosím prihláste.'
login_manager.login_message_category = 'error'

# ── Počítadlo neúspešných pokusov (v session) ────────────────────────────────
LOGIN_MAX_POKUSY  = 4
LOGIN_LOCKOUT_MIN = 15

def get_ip():
    return request.headers.get('X-Forwarded-For', request.remote_addr or 'unknown').split(',')[0].strip()

@app.errorhandler(CSRFError)
def csrf_error(e):
    flash('Platnosť formulára vypršala. Skúste znova.', 'error')
    return redirect(request.referrer or url_for('login'))

def now_sk():
    return datetime.now(timezone.utc).astimezone(timezone(timedelta(hours=2))).replace(tzinfo=None)

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

with app.app_context():
    db.create_all()
    # Migrácie — pridaj stĺpce ak neexistujú (bezpečné opakovane spustiť)
    from sqlalchemy import text
    _migrations = [
        "ALTER TABLE cinnosti ADD COLUMN IF NOT EXISTS poznamka TEXT DEFAULT ''",
        "ALTER TABLE dovolenky ADD COLUMN IF NOT EXISTS poznamka VARCHAR(200) DEFAULT ''",
        "ALTER TABLE predefinovane_cinnosti ADD COLUMN IF NOT EXISTS aktivna BOOLEAN DEFAULT TRUE",
        # ziadosti_o_zmenu vytvorí db.create_all() ako nová tabuľka
    ]
    for sql in _migrations:
        try:
            db.session.execute(text(sql))
            db.session.commit()
        except Exception:
            db.session.rollback()


# ── AUTH ──────────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    if current_user.is_authenticated:
        return redirect(url_for('dashboard'))
    return redirect(url_for('login'))

@app.route('/registracia', methods=['GET', 'POST'])
def registracia():
    if current_user.is_authenticated:
        return redirect(url_for('dashboard'))
    # Ak už existuje aspoň jeden user, vyžadujeme registračný kód
    prvy_user = User.query.count() == 0
    if request.method == 'POST':
        meno   = request.form.get('meno', '').strip()
        email  = request.form.get('email', '').strip().lower()
        heslo  = request.form.get('heslo', '')
        heslo2 = request.form.get('heslo2', '')
        reg_kod = request.form.get('reg_kod', '').strip()
        if not meno or not email or not heslo:
            flash('Vyplňte všetky polia.', 'error')
            return render_template('registracia.html', prvy_user=prvy_user)
        if heslo != heslo2:
            flash('Heslá sa nezhodujú.', 'error')
            return render_template('registracia.html', prvy_user=prvy_user)
        if len(heslo) < 6:
            flash('Heslo musí mať aspoň 6 znakov.', 'error')
            return render_template('registracia.html', prvy_user=prvy_user)
        # Kontrola registračného kódu (nie pre prvého používateľa)
        if not prvy_user:
            ulozeny_kod = Nastavenie.get('reg_kod', '')
            if not ulozeny_kod:
                flash('Registrácia je momentálne zakázaná. Kontaktujte správcu.', 'error')
                return render_template('registracia.html', prvy_user=prvy_user)
            if reg_kod != ulozeny_kod:
                flash('Nesprávny registračný kód organizácie.', 'error')
                return render_template('registracia.html', prvy_user=prvy_user)
        if User.query.filter_by(email=email).first():
            flash('Email už existuje.', 'error')
            return render_template('registracia.html', prvy_user=prvy_user)
        hash_hesla = bcrypt.hashpw(heslo.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
        is_admin = prvy_user
        user = User(meno=meno, email=email, heslo=hash_hesla, is_admin=is_admin)
        db.session.add(user)
        db.session.commit()
        flash('Registrácia úspešná! Prihláste sa.', 'success')
        return redirect(url_for('login'))
    return render_template('registracia.html', prvy_user=prvy_user)

@app.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('dashboard'))

    # Inicializácia session počítadla
    if 'login_pokusy' not in session:
        session['login_pokusy'] = 0
        session['login_lockout_do'] = None

    # Kontrola lockout
    lockout_do = session.get('login_lockout_do')
    if lockout_do:
        lockout_dt = datetime.fromisoformat(lockout_do)
        zostatok = (lockout_dt - now_sk()).total_seconds()
        if zostatok > 0:
            minuty = int(zostatok // 60) + 1
            flash(f'Príliš veľa neúspešných pokusov. Skúste znova o {minuty} min.', 'error')
            return render_template('login.html', lockout=True, lockout_minuty=minuty)
        else:
            # Lockout vypršal — reset
            session['login_pokusy'] = 0
            session['login_lockout_do'] = None

    if request.method == 'POST':
        email = request.form.get('email', '').strip().lower()
        heslo = request.form.get('heslo', '')
        user  = User.query.filter_by(email=email).first()

        if user and bcrypt.checkpw(heslo.encode('utf-8'), user.heslo.encode('utf-8')):
            # Úspešné prihlásenie — reset počítadla
            session['login_pokusy'] = 0
            session['login_lockout_do'] = None
            login_user(user, remember=True)
            # Log
            log = LoginLog(user_id=user.id, email=email, akcia='login_ok',
                           ip=get_ip(), created_at=now_sk())
            db.session.add(log)
            db.session.commit()
            return redirect(url_for('dashboard'))
        else:
            # Neúspešný pokus
            session['login_pokusy'] = session.get('login_pokusy', 0) + 1
            pokusy = session['login_pokusy']
            zostatok = LOGIN_MAX_POKUSY - pokusy
            # Log neúspešného pokusu
            uid = user.id if user else None
            log = LoginLog(user_id=uid, email=email, akcia='login_fail',
                           ip=get_ip(), created_at=now_sk())
            db.session.add(log)
            db.session.commit()

            if zostatok <= 0:
                lockout_cas = now_sk() + timedelta(minutes=LOGIN_LOCKOUT_MIN)
                session['login_lockout_do'] = lockout_cas.isoformat()
                flash(f'Príliš veľa neúspešných pokusov. Účet zablokovaný na {LOGIN_LOCKOUT_MIN} minút.', 'error')
                return render_template('login.html', lockout=True, lockout_minuty=LOGIN_LOCKOUT_MIN)
            else:
                flash(f'Nesprávny email alebo heslo. Zostáva {zostatok} pokus{"y" if zostatok > 1 else ""}.', 'error')
                return render_template('login.html', zostatok_pokusov=zostatok)

    pokusy = session.get('login_pokusy', 0)
    zostatok = max(0, LOGIN_MAX_POKUSY - pokusy)
    return render_template('login.html', zostatok_pokusov=zostatok if pokusy > 0 else None)

@app.route('/logout')
@login_required
def logout():
    email = current_user.email
    uid   = current_user.id
    logout_user()
    log = LoginLog(user_id=uid, email=email, akcia='logout',
                   ip=get_ip(), created_at=now_sk())
    db.session.add(log)
    db.session.commit()
    return redirect(url_for('login'))

@app.route('/zabudnute-heslo', methods=['GET', 'POST'])
def zabudnute_heslo():
    if current_user.is_authenticated:
        return redirect(url_for('dashboard'))
    if request.method == 'POST':
        email      = request.form.get('email', '').strip().lower()
        reset_kod  = request.form.get('reset_kod', '').strip()
        heslo      = request.form.get('heslo', '')
        heslo2     = request.form.get('heslo2', '')
        # Validácia
        ulozeny_kod = Nastavenie.get('reset_kod', '')
        if not ulozeny_kod:
            flash('Reset hesla nie je povolený. Kontaktujte správcu.', 'error')
            return render_template('zabudnute_heslo.html')
        if reset_kod != ulozeny_kod:
            flash('Nesprávny reset kód.', 'error')
            return render_template('zabudnute_heslo.html')
        user = User.query.filter_by(email=email).first()
        if not user:
            flash('Email nenájdený.', 'error')
            return render_template('zabudnute_heslo.html')
        if heslo != heslo2:
            flash('Heslá sa nezhodujú.', 'error')
            return render_template('zabudnute_heslo.html')
        if len(heslo) < 6:
            flash('Heslo musí mať aspoň 6 znakov.', 'error')
            return render_template('zabudnute_heslo.html')
        user.heslo = bcrypt.hashpw(heslo.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
        db.session.commit()
        flash('Heslo bolo zmenené. Prihláste sa.', 'success')
        return redirect(url_for('login'))
    return render_template('zabudnute_heslo.html')


# ── DASHBOARD ─────────────────────────────────────────────────────────────────

@app.route('/dashboard')
@login_required
def dashboard():
    aktivny = Zaznam.query.filter_by(user_id=current_user.id, cas_stop=None)\
                          .order_by(Zaznam.cas_start.desc()).first()
    klienti = Klient.query.filter_by(aktivny=True)\
                          .order_by(Klient.nazov).all()
    posledne = Zaznam.query.filter_by(user_id=current_user.id)\
                           .filter(Zaznam.cas_stop.isnot(None))\
                           .order_by(Zaznam.cas_start.desc()).limit(5).all()
    moje_ulohy = Uloha.query.filter_by(prijimatel_id=current_user.id, hotovo=False)\
                             .order_by(Uloha.urgencia.desc(), Uloha.created_at.desc()).all()
    pred_cinnosti = PredefinovanaCinnost.query.filter_by(aktivna=True)\
                                              .order_by(PredefinovanaCinnost.nazov).all()
    return render_template('dashboard.html', aktivny=aktivny, klienti=klienti,
                           posledne=posledne, moje_ulohy=moje_ulohy,
                           pred_cinnosti=pred_cinnosti)


# ── ČASOVAČ ───────────────────────────────────────────────────────────────────

@app.route('/start', methods=['POST'])
@login_required
def start_cas():
    if Zaznam.query.filter_by(user_id=current_user.id, cas_stop=None).first():
        flash('Máte aktívny časovač. Najprv ho zastavte.', 'error')
        return redirect(url_for('dashboard'))
    klient_id = request.form.get('klient_id')
    klient = Klient.query.filter_by(id=int(klient_id)).first()
    if not klient:
        flash('Vyberte platného klienta.', 'error')
        return redirect(url_for('dashboard'))
    now = now_sk()
    zaznam = Zaznam(user_id=current_user.id, klient_id=klient.id, datum=now.date(), cas_start=now)
    db.session.add(zaznam)
    db.session.commit()
    flash(f'Časovač spustený — {klient.nazov}', 'success')
    return redirect(url_for('dashboard'))

@app.route('/zmenit-klienta', methods=['POST'])
@login_required
def zmenit_klienta():
    zaznam_id      = request.form.get('zaznam_id')
    novy_klient_id = request.form.get('novy_klient_id')
    popis          = request.form.get('popis', '').strip()
    typ            = request.form.get('typ', '').strip()

    # Vyžaduje sa len popis a typ pre STARÉHO klienta (čo ste robili doteraz)
    if not popis:
        flash('Vyplňte popis činnosti pre aktuálneho klienta.', 'error')
        return redirect(url_for('dashboard'))
    if typ not in ('bezna', 'extra'):
        flash('Vyberte typ činnosti pre aktuálneho klienta (Bežná / Extra).', 'error')
        return redirect(url_for('dashboard'))

    zaznam = Zaznam.query.filter_by(id=int(zaznam_id), user_id=current_user.id, cas_stop=None).first()
    novy_klient = Klient.query.filter_by(id=int(novy_klient_id)).first()
    if not zaznam or not novy_klient:
        flash('Chyba — záznam alebo klient nenájdený.', 'error')
        return redirect(url_for('dashboard'))
    if novy_klient.id == zaznam.klient_id:
        flash('Vyberte iného klienta.', 'error')
        return redirect(url_for('dashboard'))

    now = now_sk()

    # Uzavri otvorenú činnosť starého klienta
    posledna = Cinnost.query.filter_by(zaznam_id=zaznam.id, cas_do=None).first()
    if posledna:
        posledna.cas_do = now

    # Pridaj záverečnú činnosť pre starého klienta (čo ste robili)
    cinnost = Cinnost(zaznam_id=zaznam.id, popis=popis, typ=typ, cas_od=now, cas_do=now)
    db.session.add(cinnost)

    # Stopni záznam starého klienta
    zaznam.cas_stop = now

    # Spusti nový záznam pre nového klienta — BEZ požiadavky na popis/typ
    # (ten sa vyplní až pri rozdelení alebo zastavení)
    novy_zaznam = Zaznam(user_id=current_user.id, klient_id=novy_klient.id,
                         datum=now.date(), cas_start=now)
    db.session.add(novy_zaznam)
    db.session.commit()
    flash(f'Prepnuté na {novy_klient.nazov} — časovač beží.', 'success')
    return redirect(url_for('dashboard'))

@app.route('/stop', methods=['POST'])
@login_required
def stop_cas():
    zaznam_id = request.form.get('zaznam_id')
    zaznam = Zaznam.query.filter_by(id=int(zaznam_id), user_id=current_user.id, cas_stop=None).first()
    if not zaznam:
        flash('Záznam nenájdený.', 'error')
        return redirect(url_for('dashboard'))
    now = now_sk()

    # Ak prišiel popis_stop (zastavenie bez predchádzajúcej činnosti)
    popis_stop   = request.form.get('popis_stop', '').strip()
    typ_stop     = request.form.get('typ', '').strip()
    poznamka_cin = request.form.get('poznamka_cin', '').strip()
    if popis_stop and typ_stop in ('bezna', 'extra'):
        # Uzavri prípadnú otvorenú činnosť
        posledna = Cinnost.query.filter_by(zaznam_id=zaznam.id, cas_do=None).first()
        if posledna:
            posledna.cas_do = now
        # Pridaj novú činnosť
        cinnost = Cinnost(zaznam_id=zaznam.id, popis=popis_stop, typ=typ_stop,
                          poznamka=poznamka_cin, cas_od=zaznam.cas_start, cas_do=now)
        db.session.add(cinnost)
    else:
        # Uzavri poslednú otvorenú činnosť
        posledna = Cinnost.query.filter_by(zaznam_id=zaznam.id, cas_do=None).first()
        if posledna:
            posledna.cas_do = now

    zaznam.cas_stop = now
    poznamka = request.form.get('poznamka', '').strip()
    if poznamka:
        zaznam.poznamka = poznamka
    db.session.commit()
    flash(f'Práca ukončená — {zaznam.trvanie_format}', 'success')
    return redirect(url_for('dashboard'))

@app.route('/cinnost/pridat', methods=['POST'])
@login_required
def pridat_cinnost():
    zaznam_id   = request.form.get('zaznam_id')
    popis       = request.form.get('popis', '').strip()
    typ         = request.form.get('typ', '').strip()
    redirect_to = request.form.get('redirect', 'dashboard')
    if not popis:
        flash('Zadajte popis činnosti.', 'error')
        return redirect(url_for(redirect_to))
    if typ not in ('bezna', 'extra'):
        flash('Vyberte typ činnosti (Bežná / Extra).', 'error')
        return redirect(url_for(redirect_to))
    poznamka_cin = request.form.get('poznamka_cin', '').strip()
    zaznam = Zaznam.query.filter_by(id=int(zaznam_id), user_id=current_user.id).first()
    if not zaznam:
        flash('Záznam nenájdený.', 'error')
        return redirect(url_for(redirect_to))
    now = now_sk()
    posledna = Cinnost.query.filter_by(zaznam_id=zaznam.id, cas_do=None).first()
    if posledna:
        posledna.cas_do = now
    cinnost = Cinnost(
        zaznam_id=zaznam.id,
        popis=popis,
        typ=typ,
        poznamka=poznamka_cin,
        cas_od=now,
        cas_do=now if zaznam.cas_stop is not None else None
    )
    db.session.add(cinnost)
    db.session.commit()
    flash('Činnosť pridaná.', 'success')
    return redirect(url_for(redirect_to))

@app.route('/zaznam/<int:zaznam_id>/poznamka', methods=['POST'])
@login_required
def ulozit_poznamku(zaznam_id):
    zaznam = Zaznam.query.filter_by(id=int(zaznam_id), user_id=current_user.id).first_or_404()
    zaznam.poznamka = request.form.get('poznamka', '').strip()
    db.session.commit()
    flash('Poznámka uložená.', 'success')
    return redirect(url_for('vykaz'))


# ── KLIENTI ───────────────────────────────────────────────────────────────────

@app.route('/klienti')
@login_required
def klienti():
    zoznam = Klient.query.order_by(Klient.nazov).all()
    return render_template('klienti.html', klienti=zoznam)

@app.route('/klient/pridat', methods=['POST'])
@login_required
def pridat_klienta():
    nazov = request.form.get('nazov', '').strip()
    if not nazov:
        flash('Zadajte názov klienta.', 'error')
        return redirect(url_for('klienti'))
    if Klient.query.filter_by(nazov=nazov).first():
        flash('Klient s týmto názvom už existuje.', 'error')
        return redirect(url_for('klienti'))
    db.session.add(Klient(nazov=nazov, user_id=current_user.id))
    db.session.commit()
    flash(f'Klient „{nazov}" pridaný.', 'success')
    return redirect(url_for('klienti'))

@app.route('/klient/<int:klient_id>/upravit', methods=['POST'])
@login_required
def upravit_klienta(klient_id):
    klient = Klient.query.get_or_404(klient_id)
    novy_nazov = request.form.get('nazov', '').strip()
    if not novy_nazov:
        flash('Názov nemôže byť prázdny.', 'error')
        return redirect(url_for('klienti'))
    klient.nazov = novy_nazov
    db.session.commit()
    flash(f'Klient premenovaný na „{novy_nazov}".', 'success')
    return redirect(url_for('klienti'))

@app.route('/klient/<int:klient_id>/archivovat', methods=['POST'])
@login_required
def archivovat_klienta(klient_id):
    klient = Klient.query.get_or_404(klient_id)
    klient.aktivny = not klient.aktivny
    db.session.commit()
    flash('Klient ' + ('aktivovaný.' if klient.aktivny else 'archivovaný.'), 'success')
    return redirect(url_for('klienti'))


# ── VÝKAZ ─────────────────────────────────────────────────────────────────────

def _filter_zaznamy(user_id, datum_od, datum_do, klient_id):
    q = Zaznam.query.filter_by(user_id=user_id).filter(Zaznam.cas_stop.isnot(None))
    if datum_od:
        try:
            q = q.filter(Zaznam.datum >= datetime.strptime(datum_od, '%Y-%m-%d').date())
        except: pass
    if datum_do:
        try:
            q = q.filter(Zaznam.datum <= datetime.strptime(datum_do, '%Y-%m-%d').date())
        except: pass
    if klient_id:
        q = q.filter(Zaznam.klient_id == int(klient_id))
    return q

@app.route('/vykaz')
@login_required
def vykaz():
    from datetime import date
    datum_od  = request.args.get('datum_od', '')
    datum_do  = request.args.get('datum_do', '')
    klient_id = request.args.get('klient_id', '')
    zaznamy   = _filter_zaznamy(current_user.id, datum_od, datum_do, klient_id)\
                    .order_by(Zaznam.datum.desc(), Zaznam.cas_start.desc()).all()
    klienti   = Klient.query.order_by(Klient.nazov).all()
    celk_sek  = sum(z.trvanie_sekundy or 0 for z in zaznamy)
    celkovy   = f'{celk_sek // 3600:02d}:{(celk_sek % 3600) // 60:02d}'

    # Dovolenky — všetky záznamy používateľa
    dovolenky = Dovolenka.query.filter_by(user_id=current_user.id)\
                               .order_by(Dovolenka.datum_od.desc()).all()

    # Dni s aktívnymi záznamami (dni kde beží alebo bežal časovač)
    pracovne_dni = set()
    for z in Zaznam.query.filter_by(user_id=current_user.id).all():
        pracovne_dni.add(z.datum)

    # Vypočítaj skutočné dovolenkové dni (bez pracovných dní)
    # Pre každú dovolenku expanduj rozsah dátumov → zoznam dní, filtruj pracovné dni
    from datetime import timedelta as td
    dovolenkove_dni = []
    for d in dovolenky:
        cur = d.datum_od
        while cur <= d.datum_do:
            if cur not in pracovne_dni:
                dovolenkove_dni.append({'datum': cur, 'dovolenka_id': d.id, 'poznamka': d.poznamka})
            cur += td(days=1)
    # Zoradiť podľa dátumu zostupne
    dovolenkove_dni.sort(key=lambda x: x['datum'], reverse=True)

    pred_cinnosti = PredefinovanaCinnost.query.filter_by(aktivna=True)\
                                              .order_by(PredefinovanaCinnost.nazov).all()
    return render_template('vykaz.html', zaznamy=zaznamy, klienti=klienti,
                           celkovy=celkovy, datum_od=datum_od, datum_do=datum_do,
                           klient_id=klient_id, dovolenky=dovolenky,
                           dovolenkove_dni=dovolenkove_dni,
                           pred_cinnosti=pred_cinnosti)

@app.route('/vykaz/export')
@login_required
def export_excel():
    datum_od  = request.args.get('datum_od', '')
    datum_do  = request.args.get('datum_do', '')
    klient_id = request.args.get('klient_id', '')
    zaznamy   = _filter_zaznamy(current_user.id, datum_od, datum_do, klient_id)\
                    .order_by(Zaznam.datum.asc(), Zaznam.cas_start.asc()).all()
    output = _build_excel(zaznamy, current_user.meno)
    filename = f'dochadzka_{current_user.meno.replace(" ","_")}_{now_sk().strftime("%Y%m%d")}.xlsx'
    return send_file(output, as_attachment=True, download_name=filename,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# ── DOVOLENKY ─────────────────────────────────────────────────────────────────

@app.route('/vykaz/dovolenka', methods=['POST'])
@login_required
def pridat_dovolenku():
    from datetime import date
    datum_od_str = request.form.get('datum_od', '').strip()
    datum_do_str = request.form.get('datum_do', '').strip()
    poznamka = request.form.get('poznamka', '').strip()
    try:
        datum_od = date.fromisoformat(datum_od_str)
        datum_do = date.fromisoformat(datum_do_str)
    except ValueError:
        flash('Neplatné dátumy dovolenky.', 'error')
        return redirect(url_for('vykaz'))
    if datum_od > datum_do:
        flash('Dátum začiatku musí byť pred dátumom konca.', 'error')
        return redirect(url_for('vykaz'))
    d = Dovolenka(user_id=current_user.id, datum_od=datum_od, datum_do=datum_do,
                  poznamka=poznamka, created_at=now_sk())
    db.session.add(d)
    db.session.commit()
    flash(f'Dovolenka {datum_od.strftime("%d.%m.%Y")} – {datum_do.strftime("%d.%m.%Y")} pridaná.', 'success')
    return redirect(url_for('vykaz'))

@app.route('/vykaz/dovolenka/<int:dov_id>/zmazat', methods=['POST'])
@login_required
def zmazat_dovolenku(dov_id):
    d = Dovolenka.query.filter_by(id=dov_id, user_id=current_user.id).first_or_404()
    db.session.delete(d)
    db.session.commit()
    flash('Dovolenka vymazaná.', 'success')
    return redirect(url_for('vykaz'))


# ── PREDEFINOVANÉ ČINNOSTI ────────────────────────────────────────────────────

@app.route('/cinnosti/predefinovane')
@login_required
def predefinovane_cinnosti():
    cinnosti = PredefinovanaCinnost.query.filter_by(aktivna=True)\
                                         .order_by(PredefinovanaCinnost.nazov).all()
    return jsonify([{'id': c.id, 'nazov': c.nazov, 'typ': c.typ} for c in cinnosti])

@app.route('/cinnosti/predefinovane/pridaj', methods=['POST'])
@login_required
def pridat_predefinovanu():
    nazov = request.form.get('nazov', '').strip()
    typ   = request.form.get('typ', 'bezna').strip()
    if not nazov:
        flash('Zadajte názov činnosti.', 'error')
        return redirect(request.referrer or url_for('dashboard'))
    if typ not in ('bezna', 'extra'):
        typ = 'bezna'
    # Kontrola duplicity
    existuje = PredefinovanaCinnost.query.filter(
        PredefinovanaCinnost.nazov.ilike(nazov),
        PredefinovanaCinnost.aktivna == True
    ).first()
    if existuje:
        flash(f'Činnosť „{nazov}" už existuje.', 'error')
        return redirect(request.referrer or url_for('dashboard'))
    c = PredefinovanaCinnost(nazov=nazov, typ=typ, created_by=current_user.id,
                              created_at=now_sk())
    db.session.add(c)
    db.session.commit()
    flash(f'Činnosť „{nazov}" pridaná.', 'success')
    return redirect(request.referrer or url_for('dashboard'))

@app.route('/cinnosti/predefinovane/<int:cin_id>/zmazat', methods=['POST'])
@login_required
def zmazat_predefinovanu(cin_id):
    c = PredefinovanaCinnost.query.get_or_404(cin_id)
    # Môže mazať iba správca alebo autor
    if not current_user.is_admin and c.created_by != current_user.id:
        flash('Prístup zamietnutý.', 'error')
        return redirect(request.referrer or url_for('dashboard'))
    c.aktivna = False
    db.session.commit()
    flash(f'Činnosť „{c.nazov}" odstránená.', 'success')
    return redirect(request.referrer or url_for('dashboard'))


# ── ŽIADOSTI O ZMENU ──────────────────────────────────────────────────────────

@app.route('/cinnost/<int:cin_id>/ziadost', methods=['POST'])
@login_required
def poslat_ziadost(cin_id):
    """Zamestnanec odošle žiadosť o zmenu konkrétnej činnosti."""
    cin = Cinnost.query.get_or_404(cin_id)
    # Overenie, že činnosť patrí aktuálnemu používateľovi
    if cin.zaznam.user_id != current_user.id:
        return jsonify({'error': 'Prístup zamietnutý'}), 403
    text = request.form.get('text', '').strip()
    if not text:
        flash('Zadajte text žiadosti.', 'error')
        return redirect(url_for('vykaz'))
    # Existujúca nevybavená žiadosť — prepíšeme
    existujuca = ZiadostOZmenu.query.filter_by(cinnost_id=cin_id, vybavena=False).first()
    if existujuca:
        existujuca.text = text
        existujuca.created_at = datetime.utcnow()
    else:
        z = ZiadostOZmenu(cinnost_id=cin_id, user_id=current_user.id, text=text)
        db.session.add(z)
    db.session.commit()
    flash('Žiadosť o zmenu odoslaná.', 'success')
    return redirect(url_for('vykaz'))

@app.route('/admin/cinnost/<int:cin_id>/vybavit_ziadost', methods=['POST'])
@login_required
def vybavit_ziadost(cin_id):
    """Správca označí žiadosť o zmenu ako vybavenú."""
    if not current_user.is_admin:
        return jsonify({'error': 'Prístup zamietnutý'}), 403
    ziadost = ZiadostOZmenu.query.filter_by(cinnost_id=cin_id, vybavena=False).first()
    if ziadost:
        ziadost.vybavena = True
        db.session.commit()
    return redirect(request.referrer or url_for('admin'))


# ── ADMIN ─────────────────────────────────────────────────────────────────────

@app.route('/admin')
@login_required
def admin():
    if not current_user.is_admin:
        flash('Prístup zamietnutý.', 'error')
        return redirect(url_for('dashboard'))
    users = User.query.order_by(User.meno).all()
    return render_template('admin.html', users=users)

@app.route('/admin/logy')
@login_required
def admin_logy():
    if not current_user.is_admin:
        flash('Prístup zamietnutý.', 'error')
        return redirect(url_for('dashboard'))
    logy = LoginLog.query.order_by(LoginLog.created_at.desc()).limit(200).all()
    return render_template('admin_logy.html', logy=logy)

@app.route('/admin/vykaz/<int:user_id>')
@login_required
def admin_vykaz(user_id):
    if not current_user.is_admin:
        flash('Prístup zamietnutý.', 'error')
        return redirect(url_for('dashboard'))
    user      = User.query.get_or_404(user_id)
    datum_od  = request.args.get('datum_od', '')
    datum_do  = request.args.get('datum_do', '')
    klient_id = request.args.get('klient_id', '')
    zaznamy   = _filter_zaznamy(user_id, datum_od, datum_do, klient_id)\
                    .order_by(Zaznam.datum.desc(), Zaznam.cas_start.desc()).all()
    klienti   = Klient.query.order_by(Klient.nazov).all()
    celk_sek  = sum(z.trvanie_sekundy or 0 for z in zaznamy)
    celkovy   = f'{celk_sek // 3600:02d}:{(celk_sek % 3600) // 60:02d}'
    return render_template('admin_vykaz.html', zamestnanec=user, zaznamy=zaznamy,
                           klienti=klienti, celkovy=celkovy,
                           datum_od=datum_od, datum_do=datum_do, klient_id=klient_id)

@app.route('/admin/export/<int:user_id>')
@login_required
def admin_export(user_id):
    if not current_user.is_admin:
        flash('Prístup zamietnutý.', 'error')
        return redirect(url_for('dashboard'))
    user      = User.query.get_or_404(user_id)
    datum_od  = request.args.get('datum_od', '')
    datum_do  = request.args.get('datum_do', '')
    klient_id = request.args.get('klient_id', '')
    zaznamy   = _filter_zaznamy(user_id, datum_od, datum_do, klient_id)\
                    .order_by(Zaznam.datum.asc(), Zaznam.cas_start.asc()).all()
    output   = _build_excel(zaznamy, user.meno, include_ziadosti=True)
    filename = f'dochadzka_{user.meno.replace(" ","_")}_{now_sk().strftime("%Y%m%d")}.xlsx'
    return send_file(output, as_attachment=True, download_name=filename,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/admin/toggle_admin/<int:user_id>', methods=['POST'])
@login_required
def toggle_admin(user_id):
    if not current_user.is_admin:
        flash('Prístup zamietnutý.', 'error')
        return redirect(url_for('dashboard'))
    user = User.query.get_or_404(user_id)
    if user.id == current_user.id:
        flash('Nemôžete zmeniť vlastnú rolu.', 'error')
        return redirect(url_for('admin'))
    user.is_admin = not user.is_admin
    db.session.commit()
    flash(f'{user.meno} — rola zmenená.', 'success')
    return redirect(url_for('admin'))


# ── ADMIN ANALYTIKA ───────────────────────────────────────────────────────────

@app.route('/admin/analyza')
@login_required
def admin_analyza():
    if not current_user.is_admin:
        flash('Prístup zamietnutý.', 'error')
        return redirect(url_for('dashboard'))

    # Načítame všetky ukončené záznamy so cinnosťami
    vsetky_zaznamy = (Zaznam.query
                     .filter(Zaznam.cas_stop.isnot(None))
                     .order_by(Zaznam.datum.asc()).all())

    wb = openpyxl.Workbook()
    navy  = '1E3A5F'; blue2 = '2E6DA4'; white = 'FFFFFF'
    green = '065F46'; green_bg = 'D1FAE5'
    orange = '92400E'; orange_bg = 'FEF3C7'
    red_bg = 'FEE2E2'; red_fg = '991B1B'
    light = 'EEF4FB'; gray = 'F1F5F9'

    hdr = lambda ws, row, col, val, bg=navy, fg=white, bold=True, size=10: _acell(ws,row,col,val,bg,fg,bold,size)

    def _acell(ws, row, col, val, bg=None, fg='000000', bold=False, size=10, align='center'):
        c = ws.cell(row=row, column=col, value=val)
        if bg:
            c.fill = PatternFill('solid', start_color=bg)
        c.font = Font(bold=bold, color=fg, name='Calibri', size=size)
        c.alignment = Alignment(horizontal=align, vertical='center', wrap_text=True)
        thin = Border(
            left=Side(style='thin', color='CCCCCC'), right=Side(style='thin', color='CCCCCC'),
            top=Side(style='thin', color='CCCCCC'),  bottom=Side(style='thin', color='CCCCCC'))
        c.border = thin
        return c

    def sek_na_hod(sek):
        return round(sek / 3600, 2)

    def format_hod(sek):
        h = sek // 3600; m = (sek % 3600) // 60
        return f'{h:02d}:{m:02d}'

    export_cas = now_sk().strftime('%d.%m.%Y %H:%M')

    # ── 1. SHEET: Prehľad podľa klientov (celkovo) ────────────────────────────
    ws1 = wb.active
    ws1.title = '1. Klienti – celkovo'

    klient_data = defaultdict(lambda: {'sek': 0, 'bezna_sek': 0, 'extra_sek': 0,
                                        'zaznamov': 0, 'cinnosti': [], 'mesiace': defaultdict(int)})
    for z in vsetky_zaznamy:
        k = z.klient.nazov if z.klient else 'Neznámy'
        mes = z.datum.strftime('%Y-%m') if z.datum else '?'
        sek = z.trvanie_sekundy or 0
        klient_data[k]['sek'] += sek
        klient_data[k]['zaznamov'] += 1
        klient_data[k]['mesiace'][mes] += sek
        for cin in z.cinnosti:
            cin_sek = int((cin.cas_do - cin.cas_od).total_seconds()) if cin.cas_do and cin.cas_od else 0
            if cin.typ == 'extra':
                klient_data[k]['extra_sek'] += cin_sek
            else:
                klient_data[k]['bezna_sek'] += cin_sek
            klient_data[k]['cinnosti'].append(cin.popis)

    ws1.merge_cells('A1:H1')
    _acell(ws1,1,1,f'Prehľad odpracovaných hodín podľa klientov  |  Export: {export_cas}',
           navy,white,True,13,'left')
    ws1.row_dimensions[1].height = 28

    hlavicky1 = ['Klient','Celkové hodiny','Bežná práca','Extra práca',
                 '% Extra','Počet záznamov','Prvý záznam','Posledný záznam']
    sirky1    = [30, 16, 16, 16, 10, 16, 14, 14]
    for ci,(h,w) in enumerate(zip(hlavicky1,sirky1),1):
        _acell(ws1,2,ci,h,navy,white,True,10)
        ws1.column_dimensions[get_column_letter(ci)].width = w
    ws1.row_dimensions[2].height = 22

    # Zoradíme podľa celkových hodín
    sorted_klienti = sorted(klient_data.items(), key=lambda x: x[1]['sek'], reverse=True)
    celk_sek_total = sum(v['sek'] for v in klient_data.values())

    for ri,(k_nazov,d) in enumerate(sorted_klienti, 3):
        bg = light if ri % 2 == 0 else white
        pct_extra = round(d['extra_sek']/d['sek']*100,1) if d['sek'] > 0 else 0
        # Prvý/posledný záznam
        daty = [z.datum for z in vsetky_zaznamy if (z.klient.nazov if z.klient else 'Neznámy') == k_nazov and z.datum]
        prvy  = min(daty).strftime('%d.%m.%Y') if daty else '—'
        posledny = max(daty).strftime('%d.%m.%Y') if daty else '—'

        _acell(ws1,ri,1,k_nazov,bg,'000000',True,10,'left')
        _acell(ws1,ri,2,format_hod(d['sek']),bg,blue2,True,10)
        _acell(ws1,ri,3,format_hod(d['bezna_sek']),bg,'065F46',False,10)
        _acell(ws1,ri,4,format_hod(d['extra_sek']),bg if pct_extra<20 else orange_bg,
               '000000' if pct_extra<20 else orange,False,10)
        # % Extra s farbením
        pct_bg = orange_bg if pct_extra >= 20 else (green_bg if pct_extra < 5 else bg)
        pct_fg = orange if pct_extra >= 20 else (green if pct_extra < 5 else '444444')
        _acell(ws1,ri,5,f'{pct_extra}%',pct_bg,pct_fg,pct_extra>=20,10)
        _acell(ws1,ri,6,d['zaznamov'],bg,'444444',False,10)
        _acell(ws1,ri,7,prvy,bg,'888888',False,9)
        _acell(ws1,ri,8,posledny,bg,'888888',False,9)
        ws1.row_dimensions[ri].height = 18

    # Súčtový riadok
    sr = len(sorted_klienti) + 3
    _acell(ws1,sr,1,'CELKOM',navy,white,True,11,'left')
    _acell(ws1,sr,2,format_hod(celk_sek_total),blue2,white,True,11)
    _acell(ws1,sr,3,format_hod(sum(v['bezna_sek'] for v in klient_data.values())),navy,white,True,11)
    _acell(ws1,sr,4,format_hod(sum(v['extra_sek'] for v in klient_data.values())),navy,white,True,11)
    ws1.row_dimensions[sr].height = 22
    ws1.freeze_panes = 'A3'

    # ── 2. SHEET: Mesačný vývoj klientov ──────────────────────────────────────
    ws2 = wb.create_sheet('2. Mesačný vývoj')

    # Zoznam všetkých mesiacov v dátach
    vsetky_mesiace = sorted(set(
        z.datum.strftime('%Y-%m') for z in vsetky_zaznamy if z.datum
    ))

    ws2.merge_cells(f'A1:{get_column_letter(len(vsetky_mesiace)+2)}1')
    _acell(ws2,1,1,f'Mesačný vývoj hodín podľa klientov  |  Export: {export_cas}',
           navy,white,True,13,'left')
    ws2.row_dimensions[1].height = 28

    _acell(ws2,2,1,'Klient',navy,white,True,10)
    ws2.column_dimensions['A'].width = 30
    _acell(ws2,2,2,'Celkom',navy,white,True,10)
    ws2.column_dimensions['B'].width = 12
    for mi,mes in enumerate(vsetky_mesiace, 3):
        rok,miac = mes.split('-')
        _acell(ws2,2,mi,f'{miac}/{rok}',navy,white,True,10)
        ws2.column_dimensions[get_column_letter(mi)].width = 11
    ws2.row_dimensions[2].height = 22

    for ri,(k_nazov,d) in enumerate(sorted_klienti, 3):
        bg = light if ri % 2 == 0 else white
        _acell(ws2,ri,1,k_nazov,bg,'000000',True,10,'left')
        _acell(ws2,ri,2,format_hod(d['sek']),bg,blue2,True,10)
        for mi,mes in enumerate(vsetky_mesiace, 3):
            sek_mes = d['mesiace'].get(mes, 0)
            cell_bg = bg
            cell_fg = '000000'
            # Zvýrazni ak je viac ako 40 hodín za mesiac (nad priemerom)
            if sek_mes > 40*3600:
                cell_bg = orange_bg; cell_fg = orange
            elif sek_mes == 0:
                cell_fg = 'CCCCCC'
            _acell(ws2,ri,mi,format_hod(sek_mes) if sek_mes else '—',cell_bg,cell_fg,sek_mes>40*3600,9)
        ws2.row_dimensions[ri].height = 18

    ws2.freeze_panes = 'C3'

    # ── 3. SHEET: Nárast/pokles hodín (MoM) ───────────────────────────────────
    ws3 = wb.create_sheet('3. Nárast hodín MoM')

    ws3.merge_cells(f'A1:{get_column_letter(len(vsetky_mesiace)+1)}1')
    _acell(ws3,1,1,f'Medzimsačný nárast odpracovaných hodín (MoM)  |  Export: {export_cas}',
           navy,white,True,13,'left')
    ws3.row_dimensions[1].height = 28

    _acell(ws3,2,1,'Klient',navy,white,True,10)
    ws3.column_dimensions['A'].width = 30
    for mi,mes in enumerate(vsetky_mesiace[1:], 2):
        rok,miac = mes.split('-')
        predch = vsetky_mesiace[vsetky_mesiace.index(mes)-1]
        pr,pm = predch.split('-')
        _acell(ws3,2,mi,f'{pm}/{pr}→{miac}/{rok}',navy,white,True,9)
        ws3.column_dimensions[get_column_letter(mi)].width = 14
    ws3.row_dimensions[2].height = 22

    for ri,(k_nazov,d) in enumerate(sorted_klienti, 3):
        bg = light if ri % 2 == 0 else white
        _acell(ws3,ri,1,k_nazov,bg,'000000',True,10,'left')
        for mi,mes in enumerate(vsetky_mesiace[1:], 2):
            predch = vsetky_mesiace[vsetky_mesiace.index(mes)-1]
            sek_teraz = d['mesiace'].get(mes, 0)
            sek_pred  = d['mesiace'].get(predch, 0)
            if sek_pred == 0 and sek_teraz == 0:
                _acell(ws3,ri,mi,'—',bg,'CCCCCC',False,9)
            elif sek_pred == 0:
                _acell(ws3,ri,mi,'Nový',green_bg,green,True,9)
            else:
                diff_pct = round((sek_teraz - sek_pred) / sek_pred * 100, 1)
                cell_bg = bg; cell_fg = '000000'
                if diff_pct >= 30:
                    cell_bg = orange_bg; cell_fg = orange
                elif diff_pct >= 10:
                    cell_bg = 'FEF9C3'; cell_fg = '78350F'
                elif diff_pct <= -30:
                    cell_bg = red_bg; cell_fg = red_fg
                elif diff_pct <= -10:
                    cell_bg = 'FEE2E2'; cell_fg = red_fg
                elif -5 <= diff_pct <= 5:
                    cell_bg = green_bg; cell_fg = green
                znacka = '▲' if diff_pct > 0 else ('▼' if diff_pct < 0 else '=')
                _acell(ws3,ri,mi,f'{znacka} {diff_pct:+.1f}%',cell_bg,cell_fg,abs(diff_pct)>=30,9)
        ws3.row_dimensions[ri].height = 18

    ws3.freeze_panes = 'B3'

    # ── 4. SHEET: Extra práce – analýza ───────────────────────────────────────
    ws4 = wb.create_sheet('4. Extra práca')

    ws4.merge_cells('A1:F1')
    _acell(ws4,1,1,f'Analýza extra práce  |  Export: {export_cas}',navy,white,True,13,'left')
    ws4.row_dimensions[1].height = 28

    extra_prace = defaultdict(lambda: {'sek': 0, 'pocet': 0, 'klienti': defaultdict(int)})
    for z in vsetky_zaznamy:
        k = z.klient.nazov if z.klient else 'Neznámy'
        for cin in z.cinnosti:
            if cin.typ == 'extra':
                cin_sek = int((cin.cas_do - cin.cas_od).total_seconds()) if cin.cas_do and cin.cas_od else 0
                popis_key = cin.popis.strip().lower()[:80] if cin.popis else ''
                extra_prace[popis_key]['sek'] += cin_sek
                extra_prace[popis_key]['pocet'] += 1
                extra_prace[popis_key]['klienti'][k] += cin_sek

    hlavicky4 = ['Popis extra činnosti','Celkový čas','Počet výskytov','Priemerný čas','Najčastejší klient','Iní klienti']
    sirky4    = [40, 14, 14, 14, 28, 20]
    for ci,(h,w) in enumerate(zip(hlavicky4,sirky4),1):
        _acell(ws4,2,ci,h,navy,white,True,10)
        ws4.column_dimensions[get_column_letter(ci)].width = w
    ws4.row_dimensions[2].height = 22

    sorted_extra = sorted(extra_prace.items(), key=lambda x: x[1]['sek'], reverse=True)
    for ri,(popis,d) in enumerate(sorted_extra, 3):
        bg = orange_bg if ri % 2 == 0 else 'FFFBF0'
        avg_sek = d['sek'] // d['pocet'] if d['pocet'] > 0 else 0
        top_klient = max(d['klienti'].items(), key=lambda x: x[1])[0] if d['klienti'] else '—'
        ostatni = ', '.join(k for k in d['klienti'] if k != top_klient) or '—'
        _acell(ws4,ri,1,popis.capitalize(),bg,'000000',False,10,'left')
        _acell(ws4,ri,2,format_hod(d['sek']),bg,orange,True,10)
        _acell(ws4,ri,3,d['pocet'],bg,'444444',False,10)
        _acell(ws4,ri,4,format_hod(avg_sek),bg,'444444',False,10)
        _acell(ws4,ri,5,top_klient,bg,'000000',False,10,'left')
        _acell(ws4,ri,6,ostatni,bg,'888888',False,9,'left')
        ws4.row_dimensions[ri].height = 18

    ws4.freeze_panes = 'A3'

    # ── 5. SHEET: Prehľad podľa zamestnancov ──────────────────────────────────
    ws5 = wb.create_sheet('5. Zamestnanci')

    ws5.merge_cells('A1:H1')
    _acell(ws5,1,1,f'Prehľad podľa zamestnancov  |  Export: {export_cas}',navy,white,True,13,'left')
    ws5.row_dimensions[1].height = 28

    # Dáta per zamestnanec
    zam_data = defaultdict(lambda: {'sek': 0, 'bezna_sek': 0, 'extra_sek': 0,
                                     'zaznamov': 0, 'klienti': set(), 'mesiace': defaultdict(int)})
    for z in vsetky_zaznamy:
        uid = z.user_id
        meno_u = z.user.meno if z.user else f'User#{uid}'
        mes = z.datum.strftime('%Y-%m') if z.datum else '?'
        sek = z.trvanie_sekundy or 0
        zam_data[meno_u]['sek'] += sek
        zam_data[meno_u]['zaznamov'] += 1
        zam_data[meno_u]['mesiace'][mes] += sek
        if z.klient:
            zam_data[meno_u]['klienti'].add(z.klient.nazov)
        for cin in z.cinnosti:
            cin_sek = int((cin.cas_do - cin.cas_od).total_seconds()) if cin.cas_do and cin.cas_od else 0
            if cin.typ == 'extra':
                zam_data[meno_u]['extra_sek'] += cin_sek
            else:
                zam_data[meno_u]['bezna_sek'] += cin_sek

    hlavicky5 = ['Zamestnanec','Celkové hodiny','Bežná práca','Extra práca',
                 '% Extra','Počet záznamov','Počet klientov','Aktívne mesiace']
    sirky5    = [28, 16, 16, 16, 10, 14, 14, 14]
    for ci,(h,w) in enumerate(zip(hlavicky5,sirky5),1):
        _acell(ws5,2,ci,h,navy,white,True,10)
        ws5.column_dimensions[get_column_letter(ci)].width = w
    ws5.row_dimensions[2].height = 22

    sorted_zam = sorted(zam_data.items(), key=lambda x: x[1]['sek'], reverse=True)
    for ri,(meno_u,d) in enumerate(sorted_zam, 3):
        bg = light if ri % 2 == 0 else white
        pct_extra = round(d['extra_sek']/d['sek']*100,1) if d['sek'] > 0 else 0
        aktivne_mesiace = len([m for m in d['mesiace'] if d['mesiace'][m] > 0])
        pct_bg = orange_bg if pct_extra >= 20 else (green_bg if pct_extra < 5 else bg)
        pct_fg = orange if pct_extra >= 20 else (green if pct_extra < 5 else '444444')
        _acell(ws5,ri,1,meno_u,bg,'000000',True,10,'left')
        _acell(ws5,ri,2,format_hod(d['sek']),bg,blue2,True,10)
        _acell(ws5,ri,3,format_hod(d['bezna_sek']),bg,green,False,10)
        _acell(ws5,ri,4,format_hod(d['extra_sek']),bg if pct_extra<20 else orange_bg,
               '000000' if pct_extra<20 else orange,False,10)
        _acell(ws5,ri,5,f'{pct_extra}%',pct_bg,pct_fg,pct_extra>=20,10)
        _acell(ws5,ri,6,d['zaznamov'],bg,'444444',False,10)
        _acell(ws5,ri,7,len(d['klienti']),bg,'444444',False,10)
        _acell(ws5,ri,8,aktivne_mesiace,bg,'444444',False,10)
        ws5.row_dimensions[ri].height = 18

    ws5.row_dimensions[2].height = 22
    ws5.freeze_panes = 'A3'

    # ── 6. SHEET: Zamestnanec × Klient krížová tabuľka ────────────────────────
    ws6 = wb.create_sheet('6. Zam × Klient')

    ws6.merge_cells(f'A1:{get_column_letter(len(sorted_klienti)+2)}1')
    _acell(ws6,1,1,f'Hodiny: Zamestnanec × Klient  |  Export: {export_cas}',navy,white,True,13,'left')
    ws6.row_dimensions[1].height = 28

    _acell(ws6,2,1,'Zamestnanec',navy,white,True,10)
    ws6.column_dimensions['A'].width = 28
    _acell(ws6,2,2,'Celkom',navy,white,True,10)
    ws6.column_dimensions['B'].width = 12
    for ci,(k_nazov,_) in enumerate(sorted_klienti, 3):
        _acell(ws6,2,ci,k_nazov,navy,white,True,9)
        ws6.column_dimensions[get_column_letter(ci)].width = 14

    # Kríž: zamestnanec × klient
    zam_klient = defaultdict(lambda: defaultdict(int))
    for z in vsetky_zaznamy:
        meno_u = z.user.meno if z.user else f'User#{z.user_id}'
        k = z.klient.nazov if z.klient else 'Neznámy'
        zam_klient[meno_u][k] += z.trvanie_sekundy or 0

    for ri,(meno_u,_) in enumerate(sorted_zam, 3):
        bg = light if ri % 2 == 0 else white
        celkom_u = zam_data[meno_u]['sek']
        _acell(ws6,ri,1,meno_u,bg,'000000',True,10,'left')
        _acell(ws6,ri,2,format_hod(celkom_u),bg,blue2,True,10)
        for ci,(k_nazov,_) in enumerate(sorted_klienti, 3):
            sek_uk = zam_klient[meno_u].get(k_nazov, 0)
            cell_bg = bg
            if sek_uk > 80*3600: cell_bg = orange_bg
            elif sek_uk > 40*3600: cell_bg = 'FEF9C3'
            elif sek_uk == 0: cell_bg = gray
            _acell(ws6,ri,ci,format_hod(sek_uk) if sek_uk else '—',cell_bg,
                   orange if sek_uk>80*3600 else ('000000' if sek_uk else 'BBBBBB'),
                   sek_uk>80*3600,9)
        ws6.row_dimensions[ri].height = 18

    ws6.row_dimensions[2].height = 22
    ws6.freeze_panes = 'C3'

    # ── 7. SHEET: Fakturačný prehľad (bežná vs extra per klient/mesiac) ────────
    ws7 = wb.create_sheet('7. Fakturácia')

    ws7.merge_cells('A1:G1')
    _acell(ws7,1,1,f'Fakturačný prehľad (bežná vs. extra práca)  |  Export: {export_cas}',
           navy,white,True,13,'left')
    ws7.row_dimensions[1].height = 28

    hlavicky7 = ['Klient','Mesiac','Bežná práca (hod)','Extra práca (hod)',
                 'Celkom (hod)','% Extra','Zamestnanci']
    sirky7    = [30, 10, 18, 18, 14, 10, 36]
    for ci,(h,w) in enumerate(zip(hlavicky7,sirky7),1):
        _acell(ws7,2,ci,h,navy,white,True,10)
        ws7.column_dimensions[get_column_letter(ci)].width = w
    ws7.row_dimensions[2].height = 22

    # Bezná a extra per klient+mesiac
    fakt_data = defaultdict(lambda: defaultdict(lambda: {'bezna': 0, 'extra': 0, 'zam': set()}))
    for z in vsetky_zaznamy:
        k = z.klient.nazov if z.klient else 'Neznámy'
        mes = z.datum.strftime('%Y-%m') if z.datum else '?'
        meno_u = z.user.meno if z.user else f'User#{z.user_id}'
        fakt_data[k][mes]['zam'].add(meno_u)
        for cin in z.cinnosti:
            cin_sek = int((cin.cas_do - cin.cas_od).total_seconds()) if cin.cas_do and cin.cas_od else 0
            if cin.typ == 'extra':
                fakt_data[k][mes]['extra'] += cin_sek
            else:
                fakt_data[k][mes]['bezna'] += cin_sek

    ri7 = 3
    for k_nazov,_ in sorted_klienti:
        mesiace_k = sorted(fakt_data[k_nazov].keys())
        for mes in mesiace_k:
            d = fakt_data[k_nazov][mes]
            celkom_sek = d['bezna'] + d['extra']
            pct_extra = round(d['extra']/celkom_sek*100,1) if celkom_sek > 0 else 0
            rok,miac = mes.split('-')
            bg = light if ri7 % 2 == 0 else white
            pct_bg = orange_bg if pct_extra >= 25 else bg
            pct_fg = orange if pct_extra >= 25 else '444444'
            _acell(ws7,ri7,1,k_nazov,bg,'000000',True,10,'left')
            _acell(ws7,ri7,2,f'{miac}/{rok}',bg,'444444',False,10)
            _acell(ws7,ri7,3,sek_na_hod(d['bezna']),bg,green,False,10)
            _acell(ws7,ri7,4,sek_na_hod(d['extra']),pct_bg,pct_fg,False,10)
            _acell(ws7,ri7,5,sek_na_hod(celkom_sek),bg,blue2,True,10)
            _acell(ws7,ri7,6,f'{pct_extra}%',pct_bg,pct_fg,pct_extra>=25,10)
            _acell(ws7,ri7,7,', '.join(sorted(d['zam'])),bg,'555555',False,9,'left')
            ws7.row_dimensions[ri7].height = 18
            ri7 += 1

    ws7.freeze_panes = 'A3'

    # ── Uloženie ──────────────────────────────────────────────────────────────
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    filename = f'analyza_dochadzka_{now_sk().strftime("%Y%m%d_%H%M")}.xlsx'
    return send_file(output, as_attachment=True, download_name=filename,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# ── NASTAVENIA ────────────────────────────────────────────────────────────────

@app.route('/nastavenia', methods=['GET', 'POST'])
@login_required
def nastavenia():
    if not current_user.is_admin:
        flash('Prístup zamietnutý. Nastavenia môže meniť iba správca.', 'error')
        return redirect(url_for('dashboard'))
    if request.method == 'POST':
        reg_kod   = request.form.get('reg_kod', '').strip()
        reset_kod = request.form.get('reset_kod', '').strip()
        Nastavenie.set('reg_kod', reg_kod)
        Nastavenie.set('reset_kod', reset_kod)
        flash('Nastavenia boli uložené.', 'success')
        return redirect(url_for('nastavenia'))
    aktualny_kod  = Nastavenie.get('reg_kod', '')
    aktualny_reset = Nastavenie.get('reset_kod', '')
    return render_template('nastavenia.html', aktualny_kod=aktualny_kod, aktualny_reset=aktualny_reset)


# ── ÚLOHY ─────────────────────────────────────────────────────────────────────

@app.route('/ulohy')
@login_required
def ulohy():
    odoslane = Uloha.query.filter_by(odosielatel_id=current_user.id)\
                          .order_by(Uloha.hotovo.asc(), Uloha.urgencia.desc(), Uloha.created_at.desc()).all()
    prijate  = Uloha.query.filter_by(prijimatel_id=current_user.id)\
                          .order_by(Uloha.hotovo.asc(), Uloha.urgencia.desc(), Uloha.created_at.desc()).all()
    users    = User.query.filter(User.id != current_user.id).order_by(User.meno).all()
    return render_template('ulohy.html', odoslane=odoslane, prijate=prijate, users=users)

@app.route('/uloha/pridat', methods=['POST'])
@login_required
def pridat_ulohu():
    prijimatel_id = request.form.get('prijimatel_id')
    text          = request.form.get('text', '').strip()
    urgencia      = int(request.form.get('urgencia', 2))
    if not text or not prijimatel_id:
        flash('Vyplňte všetky polia.', 'error')
        return redirect(url_for('ulohy'))
    prijimatel = User.query.get(prijimatel_id)
    if not prijimatel:
        flash('Príjemca nenájdený.', 'error')
        return redirect(url_for('ulohy'))
    uloha = Uloha(
        odosielatel_id=current_user.id,
        prijimatel_id=int(prijimatel_id),
        text=text,
        urgencia=urgencia,
        created_at=now_sk()
    )
    db.session.add(uloha)
    db.session.commit()
    flash(f'Úloha odoslaná pre {prijimatel.meno}.', 'success')
    return redirect(url_for('ulohy'))

@app.route('/uloha/<int:uloha_id>/hotovo', methods=['POST'])
@login_required
def oznacit_hotovo(uloha_id):
    uloha = Uloha.query.filter_by(id=int(uloha_id), prijimatel_id=current_user.id).first_or_404()
    uloha.hotovo    = not uloha.hotovo
    uloha.hotovo_at = now_sk() if uloha.hotovo else None
    db.session.commit()
    return redirect(url_for('ulohy'))


# ── EXCEL BUILDER ─────────────────────────────────────────────────────────────

def _build_excel(zaznamy, meno, include_ziadosti=False):
    # Stĺpce: A=Dátum B=Klient C=Začiatok D=Koniec E=Celkový čas F=Typ G=Činnosť H=Poznámka [I=Žiadosť o zmenu]
    num_cols = 9 if include_ziadosti else 8
    last_col_letter = get_column_letter(num_cols)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Dochádzka'
    navy  = '1E3A5F'; blue2 = '2E6DA4'; white = 'FFFFFF'; light = 'EEF4FB'
    hdr_font = Font(bold=True, color=white, name='Calibri', size=11)
    hdr_fill = PatternFill('solid', start_color=navy)
    alt_fill = PatternFill('solid', start_color=light)
    center   = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left     = Alignment(horizontal='left',   vertical='center', wrap_text=True)
    thin = lambda: Border(
        left=Side(style='thin', color='CCCCCC'), right=Side(style='thin', color='CCCCCC'),
        top=Side(style='thin', color='CCCCCC'),  bottom=Side(style='thin', color='CCCCCC'))

    # Nadpis
    ws.merge_cells(f'A1:{last_col_letter}1')
    ws['A1'] = f'Výkaz dochádzky — {meno}'
    ws['A1'].font = Font(bold=True, size=15, color=navy, name='Calibri')
    ws['A1'].alignment = center
    ws.row_dimensions[1].height = 32
    ws.merge_cells(f'A2:{last_col_letter}2')
    ws['A2'] = f'Exportované: {now_sk().strftime("%d.%m.%Y %H:%M")}'
    ws['A2'].font = Font(size=9, color='888888', name='Calibri')
    ws['A2'].alignment = center
    ws.row_dimensions[2].height = 15

    cols   = ['Dátum', 'Klient', 'Začiatok', 'Koniec', 'Celkový čas', 'Typ', 'Činnosť', 'Poznámka']
    widths = [14,       22,       11,          11,        14,            12,    38,          30]
    if include_ziadosti:
        cols.append('Žiadosť o zmenu')
        widths.append(40)
    for ci, (h, w) in enumerate(zip(cols, widths), 1):
        c = ws.cell(row=3, column=ci, value=h)
        c.font = hdr_font; c.fill = hdr_fill
        c.alignment = center; c.border = thin()
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[3].height = 24

    # Zaznamenaj prvý dátový riadok a posledný pre sumárny vzorec
    first_data_row = 4
    row = first_data_row
    for i, z in enumerate(zaznamy):
        cinnosti = z.cinnosti if z.cinnosti else [None]
        fill = alt_fill if (i % 2 == 0) else PatternFill('solid', start_color=white)
        # Priprav hodnoty spoločné pre všetky riadky záznamu
        datum_val  = z.datum.strftime('%d.%m.%Y') if z.datum else ''
        klient_val = z.klient.nazov if z.klient else ''
        # C a D ako Excel time fraction (hh:mm:ss) — čas dňa bez dátumu
        def to_time_frac(dt):
            if not dt: return None
            return (dt.hour * 3600 + dt.minute * 60 + dt.second) / 86400
        start_frac = to_time_frac(z.cas_start)
        stop_frac  = to_time_frac(z.cas_stop)

        yellow_fill = PatternFill('solid', start_color='FFFF00')

        for ji, cin in enumerate(cinnosti):
            for ci in range(1, num_cols + 1):
                c = ws.cell(row=row, column=ci)
                c.border = thin(); c.fill = fill
                c.alignment = left; c.font = Font(name='Calibri', size=10)

            # Každý riadok — dátum, klient, čas začiatku/konca
            ws.cell(row=row, column=1).value = datum_val
            ws.cell(row=row, column=1).alignment = center
            ws.cell(row=row, column=2).value = klient_val

            # Začiatok (C) — čas s sekundami
            c3 = ws.cell(row=row, column=3)
            c3.value = start_frac
            c3.number_format = 'hh:mm:ss'
            c3.alignment = center

            # Koniec (D) — čas s sekundami
            c4 = ws.cell(row=row, column=4)
            c4.value = stop_frac
            c4.number_format = 'hh:mm:ss'
            c4.alignment = center

            # Celkový čas (E) — vzorec =D-C per riadok, žlté pozadie
            if stop_frac is not None and start_frac is not None:
                c5 = ws.cell(row=row, column=5)
                c5.value = f'=D{row}-C{row}'
                c5.number_format = '[h]:mm:ss'
                c5.alignment = center
                c5.font = Font(name='Calibri', size=10, bold=True)
                c5.fill = yellow_fill

            # Typ a činnosť
            if cin:
                typ_label = 'Extra' if getattr(cin, 'typ', 'bezna') == 'extra' else 'Bežná'
                typ_fill_color = 'FEF3C7' if getattr(cin, 'typ', 'bezna') == 'extra' else 'DBEAFE'
                typ_font_color = '92400E' if getattr(cin, 'typ', 'bezna') == 'extra' else '1E40AF'
                c6 = ws.cell(row=row, column=6, value=typ_label)
                c6.alignment = center
                c6.fill = PatternFill('solid', start_color=typ_fill_color)
                c6.font = Font(name='Calibri', size=10, bold=True, color=typ_font_color)
                ws.cell(row=row, column=7).value = cin.popis
                # Poznámka k činnosti (stĺpec H)
                ws.cell(row=row, column=8).value = getattr(cin, 'poznamka', '') or ''
                # Žiadosť o zmenu (stĺpec I) — iba v admin exporte
                if include_ziadosti and cin:
                    ziadost = ZiadostOZmenu.query.filter_by(cinnost_id=cin.id, vybavena=False).first()
                    if ziadost:
                        c9 = ws.cell(row=row, column=9)
                        c9.value = ziadost.text
                        c9.fill = PatternFill('solid', start_color='FEE2E2')  # svetločervená
                        c9.font = Font(name='Calibri', size=10, color='991B1B')
            ws.row_dimensions[row].height = 18
            row += 1

    last_data_row = row - 1

    # Sumárny riadok — celkový čas ako SUM vzorec (stĺpec E = col 5)
    row += 1
    ws.merge_cells(f'A{row}:D{row}')
    ws.cell(row=row, column=1).value = 'CELKOVÝ ČAS:'
    ws.cell(row=row, column=1).font = Font(bold=True, name='Calibri', size=11, color=white)
    ws.cell(row=row, column=1).fill = PatternFill('solid', start_color=navy)
    ws.cell(row=row, column=1).alignment = Alignment(horizontal='right', vertical='center')
    # SUM vzorec na stĺpec E (Celkový čas)
    c_sum = ws.cell(row=row, column=5)
    c_sum.value = f'=SUM(E{first_data_row}:E{last_data_row})'
    c_sum.number_format = '[h]:mm:ss'
    c_sum.font = Font(bold=True, name='Calibri', size=12, color=white)
    c_sum.fill = PatternFill('solid', start_color=blue2)
    c_sum.alignment = center
    c_sum.border = thin()
    ws.merge_cells(f'F{row}:{last_col_letter}{row}')
    ws.row_dimensions[row].height = 24
    ws.freeze_panes = 'A4'
    output = io.BytesIO(); wb.save(output); output.seek(0)
    return output


# ── ZÁLOHA / OBNOVENIE ────────────────────────────────────────────────────────

@app.route('/nastavenia/zaloha')
@login_required
def zaloha():
    if not current_user.is_admin:
        flash('Prístup zamietnutý.', 'error')
        return redirect(url_for('dashboard'))

    wb = openpyxl.Workbook()
    navy  = '1E3A5F'
    white = 'FFFFFF'
    hdr_font = Font(bold=True, color=white, name='Calibri', size=11)
    hdr_fill = PatternFill('solid', start_color=navy)

    def write_sheet(ws, headers, rows):
        for col, h in enumerate(headers, 1):
            c = ws.cell(row=1, column=col, value=h)
            c.font = hdr_font
            c.fill = hdr_fill
            c.alignment = Alignment(horizontal='center', vertical='center')
        for ri, row in enumerate(rows, 2):
            for ci, val in enumerate(row, 1):
                ws.cell(row=ri, column=ci, value=val)
        for col in ws.columns:
            max_len = max((len(str(c.value or '')) for c in col), default=8)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 50)

    # Sheet 1: Používatelia
    ws_u = wb.active
    ws_u.title = 'users'
    users = User.query.order_by(User.id).all()
    write_sheet(ws_u,
        ['id', 'meno', 'email', 'heslo_hash', 'is_admin', 'created_at'],
        [(u.id, u.meno, u.email, u.heslo, int(u.is_admin),
          u.created_at.isoformat() if u.created_at else '') for u in users])

    # Sheet 2: Klienti
    ws_k = wb.create_sheet('klienti')
    klienti = Klient.query.order_by(Klient.id).all()
    write_sheet(ws_k,
        ['id', 'nazov', 'user_id', 'aktivny', 'created_at'],
        [(k.id, k.nazov, k.user_id, int(k.aktivny),
          k.created_at.isoformat() if k.created_at else '') for k in klienti])

    # Sheet 3: Záznamy
    ws_z = wb.create_sheet('zaznamy')
    zaznamy = Zaznam.query.order_by(Zaznam.id).all()
    write_sheet(ws_z,
        ['id', 'user_id', 'klient_id', 'datum', 'cas_start', 'cas_stop', 'poznamka'],
        [(z.id, z.user_id, z.klient_id,
          z.datum.isoformat() if z.datum else '',
          z.cas_start.isoformat() if z.cas_start else '',
          z.cas_stop.isoformat() if z.cas_stop else '',
          z.poznamka or '') for z in zaznamy])

    # Sheet 4: Činnosti
    ws_c = wb.create_sheet('cinnosti')
    cinnosti = Cinnost.query.order_by(Cinnost.id).all()
    write_sheet(ws_c,
        ['id', 'zaznam_id', 'popis', 'typ', 'cas_od', 'cas_do'],
        [(c.id, c.zaznam_id, c.popis, c.typ,
          c.cas_od.isoformat() if c.cas_od else '',
          c.cas_do.isoformat() if c.cas_do else '') for c in cinnosti])

    # Sheet 5: Úlohy
    ws_ul = wb.create_sheet('ulohy')
    ulohy = Uloha.query.order_by(Uloha.id).all()
    write_sheet(ws_ul,
        ['id', 'odosielatel_id', 'prijimatel_id', 'text', 'urgencia', 'hotovo', 'hotovo_at', 'created_at'],
        [(u.id, u.odosielatel_id, u.prijimatel_id, u.text, u.urgencia,
          int(u.hotovo),
          u.hotovo_at.isoformat() if u.hotovo_at else '',
          u.created_at.isoformat() if u.created_at else '') for u in ulohy])

    # Sheet 6: Nastavenia
    ws_n = wb.create_sheet('nastavenia')
    nastavenia_rows = Nastavenie.query.order_by(Nastavenie.id).all()
    write_sheet(ws_n,
        ['id', 'kluc', 'hodnota'],
        [(n.id, n.kluc, n.hodnota) for n in nastavenia_rows])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    filename = f'zaloha_dochadzka_{now_sk().strftime("%Y%m%d_%H%M")}.xlsx'
    return send_file(output, as_attachment=True, download_name=filename,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/nastavenia/obnovit', methods=['POST'])
@login_required
def obnovit_zalohu():
    if not current_user.is_admin:
        flash('Prístup zamietnutý.', 'error')
        return redirect(url_for('dashboard'))

    subor = request.files.get('zaloha_subor')
    if not subor or not subor.filename.endswith('.xlsx'):
        flash('Vyberte platný .xlsx súbor zálohy.', 'error')
        return redirect(url_for('nastavenia'))

    try:
        wb = openpyxl.load_workbook(subor, read_only=True, data_only=True)
        required = {'users', 'klienti', 'zaznamy', 'cinnosti', 'ulohy', 'nastavenia'}
        if not required.issubset(set(wb.sheetnames)):
            flash('Súbor nie je platná záloha — chýbajú listy.', 'error')
            return redirect(url_for('nastavenia'))

        def rows_of(sheet_name):
            ws = wb[sheet_name]
            rows = list(ws.iter_rows(values_only=True))
            if len(rows) < 1:
                return []
            return rows[1:]  # preskočiť hlavičku

        def parse_dt(val):
            if not val:
                return None
            if isinstance(val, datetime):
                return val.replace(tzinfo=None)
            try:
                return datetime.fromisoformat(str(val))
            except Exception:
                return None

        def parse_date(val):
            if not val:
                return None
            from datetime import date
            if isinstance(val, date):
                return val
            try:
                return datetime.fromisoformat(str(val)).date()
            except Exception:
                return None

        # Vymaž všetko v správnom poradí (FK závislosti)
        Cinnost.query.delete()
        Uloha.query.delete()
        Zaznam.query.delete()
        Klient.query.delete()
        Nastavenie.query.delete()
        User.query.delete()
        db.session.flush()

        # Obnov používateľov
        for row in rows_of('users'):
            if not row[0]:
                continue
            u = User(id=int(row[0]), meno=str(row[1] or ''), email=str(row[2] or ''),
                     heslo=str(row[3] or ''), is_admin=bool(int(row[4] or 0)),
                     created_at=parse_dt(row[5]) or now_sk())
            db.session.add(u)
        db.session.flush()

        # Obnov klientov
        for row in rows_of('klienti'):
            if not row[0]:
                continue
            k = Klient(id=int(row[0]), nazov=str(row[1] or ''), user_id=int(row[2] or 0),
                       aktivny=bool(int(row[3] or 1)),
                       created_at=parse_dt(row[4]) or now_sk())
            db.session.add(k)
        db.session.flush()

        # Obnov záznamy
        for row in rows_of('zaznamy'):
            if not row[0]:
                continue
            z = Zaznam(id=int(row[0]), user_id=int(row[1] or 0), klient_id=int(row[2] or 0),
                       datum=parse_date(row[3]), cas_start=parse_dt(row[4]),
                       cas_stop=parse_dt(row[5]), poznamka=str(row[6] or ''))
            db.session.add(z)
        db.session.flush()

        # Obnov činnosti
        for row in rows_of('cinnosti'):
            if not row[0]:
                continue
            c = Cinnost(id=int(row[0]), zaznam_id=int(row[1] or 0),
                        popis=str(row[2] or ''), typ=str(row[3] or 'bezna'),
                        cas_od=parse_dt(row[4]), cas_do=parse_dt(row[5]))
            db.session.add(c)
        db.session.flush()

        # Obnov úlohy
        for row in rows_of('ulohy'):
            if not row[0]:
                continue
            ul = Uloha(id=int(row[0]), odosielatel_id=int(row[1] or 0),
                       prijimatel_id=int(row[2] or 0), text=str(row[3] or ''),
                       urgencia=int(row[4] or 2), hotovo=bool(int(row[5] or 0)),
                       hotovo_at=parse_dt(row[6]), created_at=parse_dt(row[7]) or now_sk())
            db.session.add(ul)
        db.session.flush()

        # Obnov nastavenia
        for row in rows_of('nastavenia'):
            if not row[0]:
                continue
            n = Nastavenie(id=int(row[0]), kluc=str(row[1] or ''), hodnota=str(row[2] or ''))
            db.session.add(n)

        db.session.commit()
        flash('✅ Záloha bola úspešne obnovená. Všetky dáta sú obnovené.', 'success')

    except Exception as e:
        db.session.rollback()
        flash(f'Chyba pri obnove zálohy: {str(e)}', 'error')

    return redirect(url_for('nastavenia'))


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=False)
