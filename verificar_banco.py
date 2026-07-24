"""Diagnóstico simples do banco configurado no ambiente."""
from app import app, db, Usuario, Chamado, Financeiro, Agenda

with app.app_context():
    print('Banco:', db.engine.url.render_as_string(hide_password=True))
    print('Usuários:', Usuario.query.count())
    print('SATs:', Chamado.query.count())
    print('Financeiro:', Financeiro.query.count())
    print('Agenda:', Agenda.query.count())
