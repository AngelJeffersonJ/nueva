from functools import wraps
from flask import session, redirect, url_for, flash

def login_required(f):
    @wraps(f)
    def w(*args,**kw):
        if 'user' not in session:
            return redirect(url_for('login'))
        return f(*args,**kw)
    return w

def roles_required(*roles):
    def deco(f):
        @wraps(f)
        def w(*args,**kw):
            if session.get('role') not in roles:
                flash("Acceso denegado","danger")
                return redirect(url_for('dashboard'))
            return f(*args,**kw)
        return w
    return deco
