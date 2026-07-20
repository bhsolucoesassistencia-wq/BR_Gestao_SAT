import os, csv, io, json, uuid, hashlib, base64, socket, re, unicodedata
from datetime import datetime, date
from functools import wraps
from pathlib import Path
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file, jsonify, abort, Response
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import func, case, or_, inspect, text
from werkzeug.utils import secure_filename
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

BASE=Path(__file__).resolve().parent
UPLOAD=BASE/'static'/'uploads'; UPLOAD.mkdir(parents=True,exist_ok=True)

IMPORT_DIR=BASE/'data'/'imports'; IMPORT_DIR.mkdir(parents=True,exist_ok=True)

def norm_header(value):
    text=unicodedata.normalize('NFKD',str(value or '')).encode('ascii','ignore').decode().lower().strip()
    return re.sub(r'[^a-z0-9]+',' ',text).strip()

COLUMN_ALIASES={
 'sat':['sat','n sat','numero sat','numero do chamado','chamado','protocolo','id chamado','codigo'],
 'empreendimento':['empreendimento','condominio','obra','projeto','residencial'],
 'data_recebido':['data recebido','data recebimento','data abertura','abertura','data do chamado','data'],
 'solicitante':['solicitante','morador','cliente','nome morador','proprietario','nome cliente'],
 'unidade':['unidade','apartamento','apto','apt','torre apartamento','bloco unidade'],
 'problema':['problema','descricao','solicitacao','assunto','defeito','ocorrencia','relato'],
 'contato':['contato','telefone','celular','whatsapp','fone'],
 'status':['status','situacao','andamento'],
 'classificacao':['classificacao','procedencia','resultado'],
 'categoria':['categoria','tipo','sistema','item'],
 'data_entrega':['data entrega','entrega empreendimento','data de entrega'],
 'responsavel':['responsavel','tecnico','equipe','colaborador'],
 'observacoes':['observacoes','observacao','obs','comentarios']
}
ALIAS_LOOKUP={norm_header(a):field for field,aliases in COLUMN_ALIASES.items() for a in aliases}

def excel_date(value):
    if value is None: return ''
    if isinstance(value,(datetime,date)): return value.strftime('%Y-%m-%d')
    text=str(value).strip()
    for fmt in ('%d/%m/%Y','%Y-%m-%d','%d-%m-%Y','%d/%m/%y'):
        try: return datetime.strptime(text,fmt).strftime('%Y-%m-%d')
        except ValueError: pass
    return text

def read_import_file(file_storage):
    filename=(file_storage.filename or '').lower()
    rows=[]
    if filename.endswith('.csv'):
        raw=file_storage.read()
        text=raw.decode('utf-8-sig',errors='replace')
        sample=text[:4096]
        try: dialect=csv.Sniffer().sniff(sample,delimiters=';,\t,')
        except Exception: dialect=csv.excel; dialect.delimiter=';'
        rows=list(csv.reader(io.StringIO(text),dialect))
    elif filename.endswith(('.xlsx','.xlsm')):
        wb=load_workbook(file_storage,read_only=True,data_only=True)
        ws=wb.active
        rows=[list(r) for r in ws.iter_rows(values_only=True)]
    else:
        raise ValueError('Envie uma planilha .xlsx, .xlsm ou .csv.')
    if not rows: raise ValueError('A planilha está vazia.')
    header_idx=0; best=-1
    for i,row in enumerate(rows[:12]):
        score=sum(1 for v in row if norm_header(v) in ALIAS_LOOKUP)
        if score>best: best=score; header_idx=i
    headers=[norm_header(v) for v in rows[header_idx]]
    mapping={i:ALIAS_LOOKUP[h] for i,h in enumerate(headers) if h in ALIAS_LOOKUP}
    if 'problema' not in mapping.values() and 'sat' not in mapping.values():
        raise ValueError('Não foi possível reconhecer as colunas. Use o modelo disponível no sistema.')
    parsed=[]
    for n,row in enumerate(rows[header_idx+1:],start=header_idx+2):
        if not any(v not in (None,'') for v in row): continue
        item={'_linha':n}
        for i,field in mapping.items():
            value=row[i] if i<len(row) else ''
            item[field]=excel_date(value) if field in ('data_recebido','data_entrega') else str(value or '').strip()
        item.setdefault('data_recebido',date.today().isoformat())
        item.setdefault('status','Aberto')
        item.setdefault('categoria','')
        item['_erro']='' if (item.get('sat') or item.get('problema')) else 'Informe SAT ou descrição do problema.'
        parsed.append(item)
    return parsed, mapping
DB_URL=os.environ.get('DATABASE_URL') or f"sqlite:///{BASE/'data'/'gestao10.db'}"
if DB_URL.startswith('postgres://'): DB_URL='postgresql://'+DB_URL[len('postgres://'):]

app=Flask(__name__)
app.secret_key=os.environ.get('SECRET_KEY','troque-esta-chave-no-render-brsmartsat12')
app.config.update(SQLALCHEMY_DATABASE_URI=DB_URL,SQLALCHEMY_TRACK_MODIFICATIONS=False,MAX_CONTENT_LENGTH=40*1024*1024)
db=SQLAlchemy(app)

def ph(s): return hashlib.sha256(s.encode()).hexdigest()
def check(h,s): return h==ph(s)

class Usuario(db.Model):
    __tablename__='usuarios'; id=db.Column(db.Integer,primary_key=True); nome=db.Column(db.String(120)); email=db.Column(db.String(180),unique=True,nullable=False); senha=db.Column(db.String(128)); perfil=db.Column(db.String(30),default='Equipe'); ativo=db.Column(db.Boolean,default=True)
class Empreendimento(db.Model):
    __tablename__='empreendimentos'; id=db.Column(db.Integer,primary_key=True); nome=db.Column(db.String(150),unique=True); data_entrega=db.Column(db.String(10)); observacoes=db.Column(db.Text)
class Chamado(db.Model):
    __tablename__='chamados'; id=db.Column(db.Integer,primary_key=True); sat=db.Column(db.String(80)); empreendimento=db.Column(db.String(150)); data_recebido=db.Column(db.String(10)); solicitante=db.Column(db.String(180)); unidade=db.Column(db.String(100)); problema=db.Column(db.Text); contato=db.Column(db.String(120)); status=db.Column(db.String(80),default='Aberto'); classificacao=db.Column(db.String(100)); categoria=db.Column(db.String(220)); data_entrega=db.Column(db.String(10)); analise_garantia=db.Column(db.String(120)); fundamentacao=db.Column(db.Text); responsavel=db.Column(db.String(180)); observacoes=db.Column(db.Text); atualizado_em=db.Column(db.DateTime,default=datetime.utcnow,onupdate=datetime.utcnow)
class Garantia(db.Model):
    __tablename__='garantias'; id=db.Column(db.Integer,primary_key=True); empreendimento=db.Column(db.String(150)); item=db.Column(db.String(220)); prazo_meses=db.Column(db.Integer); classificacao_padrao=db.Column(db.String(100)); descricao=db.Column(db.Text); fonte=db.Column(db.Text)
class Foto(db.Model):
    __tablename__='fotos'; id=db.Column(db.Integer,primary_key=True); chamado_id=db.Column(db.Integer,db.ForeignKey('chamados.id')); arquivo=db.Column(db.String(255)); descricao=db.Column(db.String(255)); criado_em=db.Column(db.DateTime,default=datetime.utcnow); latitude=db.Column(db.Float); longitude=db.Column(db.Float); precisao=db.Column(db.Float); capturado_em=db.Column(db.DateTime); usuario=db.Column(db.String(180)); dispositivo=db.Column(db.String(255))
class Assinatura(db.Model):
    __tablename__='assinaturas'; id=db.Column(db.Integer,primary_key=True); chamado_id=db.Column(db.Integer,db.ForeignKey('chamados.id')); arquivo=db.Column(db.String(255)); nome_cliente=db.Column(db.String(180)); documento=db.Column(db.String(80)); observacao=db.Column(db.Text); criado_em=db.Column(db.DateTime,default=datetime.utcnow); latitude=db.Column(db.Float); longitude=db.Column(db.Float); precisao=db.Column(db.Float); assinado_em=db.Column(db.DateTime); usuario=db.Column(db.String(180)); dispositivo=db.Column(db.String(255))
class Material(db.Model):
    __tablename__='materiais'; id=db.Column(db.Integer,primary_key=True); chamado_id=db.Column(db.Integer,db.ForeignKey('chamados.id')); item=db.Column(db.String(180)); quantidade=db.Column(db.String(80)); status=db.Column(db.String(80)); observacoes=db.Column(db.Text)
class Financeiro(db.Model):
    __tablename__='financeiro'; id=db.Column(db.Integer,primary_key=True); data=db.Column(db.String(10)); tipo=db.Column(db.String(20)); descricao=db.Column(db.Text); empreendimento=db.Column(db.String(150)); categoria=db.Column(db.String(120)); valor=db.Column(db.Float,default=0); status=db.Column(db.String(50)); forma_pagamento=db.Column(db.String(50))
class Historico(db.Model):
    __tablename__='historico'; id=db.Column(db.Integer,primary_key=True); chamado_id=db.Column(db.Integer,db.ForeignKey('chamados.id')); usuario=db.Column(db.String(180)); acao=db.Column(db.String(220)); criado_em=db.Column(db.DateTime,default=datetime.utcnow)
class Auditoria(db.Model):
    __tablename__='auditoria'; id=db.Column(db.Integer,primary_key=True); modulo=db.Column(db.String(50)); registro_id=db.Column(db.Integer); usuario=db.Column(db.String(180)); acao=db.Column(db.String(80)); detalhes=db.Column(db.Text); criado_em=db.Column(db.DateTime,default=datetime.utcnow)

class Atendimento(db.Model):
    __tablename__='atendimentos'; id=db.Column(db.Integer,primary_key=True); chamado_id=db.Column(db.Integer,db.ForeignKey('chamados.id')); tipo=db.Column(db.String(30)); criado_em=db.Column(db.DateTime,default=datetime.utcnow); latitude=db.Column(db.Float); longitude=db.Column(db.Float); precisao=db.Column(db.Float); usuario=db.Column(db.String(180)); dispositivo=db.Column(db.String(255)); observacao=db.Column(db.Text)
class RelatorioTecnico(db.Model):
    __tablename__='relatorios_tecnicos'; id=db.Column(db.Integer,primary_key=True); chamado_id=db.Column(db.Integer,db.ForeignKey('chamados.id'),unique=True); diagnostico=db.Column(db.Text); servico_executado=db.Column(db.Text); testes_realizados=db.Column(db.Text); conclusao=db.Column(db.Text); atualizado_em=db.Column(db.DateTime,default=datetime.utcnow,onupdate=datetime.utcnow); usuario=db.Column(db.String(180))

def ensure_columns():
    # Migração simples para instalações que já possuem o banco das versões anteriores.
    required={
      'fotos': {'latitude':'FLOAT','longitude':'FLOAT','precisao':'FLOAT','capturado_em':'TIMESTAMP','usuario':'VARCHAR(180)','dispositivo':'VARCHAR(255)'},
      'assinaturas': {'latitude':'FLOAT','longitude':'FLOAT','precisao':'FLOAT','assinado_em':'TIMESTAMP','usuario':'VARCHAR(180)','dispositivo':'VARCHAR(255)'}
    }
    inspector=inspect(db.engine)
    tables=set(inspector.get_table_names())
    for table,cols in required.items():
        if table not in tables: continue
        existing={c['name'] for c in inspector.get_columns(table)}
        for col,typ in cols.items():
            if col not in existing:
                db.session.execute(text(f'ALTER TABLE {table} ADD COLUMN {col} {typ}'))
        db.session.commit()

def parse_location(form):
    try:
        lat=float(form.get('latitude','')); lon=float(form.get('longitude',''))
        acc=float(form.get('precisao') or 0)
        if not (-90 <= lat <= 90 and -180 <= lon <= 180): raise ValueError
        return lat,lon,acc
    except (TypeError,ValueError):
        return None

def parse_client_time(value):
    try:
        return datetime.fromisoformat((value or '').replace('Z','+00:00')).replace(tzinfo=None)
    except Exception:
        return datetime.utcnow()

def init_db():
    db.create_all()
    ensure_columns()
    if not Usuario.query.first():
        db.session.add(Usuario(nome='Administrador BR',email='admin@brsolucoes.com.br',senha=ph('123456'),perfil='Administrador'))
        db.session.commit()
    if not Chamado.query.first():
        p=BASE/'data'/'seed.json'
        if p.exists():
            data=json.loads(p.read_text(encoding='utf-8'))
            maps=[('empreendimentos',Empreendimento),('chamados',Chamado),('garantias',Garantia),('fotos',Foto),('materiais',Material),('financeiro',Financeiro)]
            for key,model in maps:
                cols={c.name for c in model.__table__.columns}
                for row in data.get(key,[]):
                    row={k:v for k,v in row.items() if k in cols and k!='id'}
                    db.session.add(model(**row))
            db.session.commit()

def login_required(fn):
    @wraps(fn)
    def w(*a,**k):
        if not session.get('user_id'): return redirect(url_for('login'))
        return fn(*a,**k)
    return w

def allow(*profiles):
    def deco(fn):
        @wraps(fn)
        def w(*a,**k):
            if session.get('perfil') not in profiles:
                flash('Seu usuário não possui permissão para acessar esta área.','erro')
                return redirect(url_for('dashboard'))
            return fn(*a,**k)
        return w
    return deco

def can_edit(): return session.get('perfil') in ('Administrador','Líder','Equipe','Técnico')
def is_admin(): return session.get('perfil')=='Administrador'

def log(cid,acao):
    db.session.add(Historico(chamado_id=cid,usuario=session.get('nome','Sistema'),acao=acao)); db.session.commit()

def audit(modulo,registro_id,acao,detalhes=''):
    db.session.add(Auditoria(modulo=modulo,registro_id=registro_id,usuario=session.get('nome','Sistema'),acao=acao,detalhes=detalhes)); db.session.commit()

def snapshot(obj, fields):
    return ' | '.join(f'{f}: {getattr(obj,f,None)}' for f in fields)

def guess_category(text):
    import re
    t=(text or '').lower(); rules=[
      ('infiltra|vazamento|umidade|mofo','Instalações hidráulicas - vedação / vazamento'),('janela|esquadria|vedação','Esquadrias de alumínio - vedação e funcionamento'),('telha|telhado|cobertura|calha','Cobertura e telhados'),('drywall|gesso','Drywall - fissuras'),('pintura|descasc|empol|tinta','Pintura interna'),('fissura|trinca|rachadura','Revestimento em argamassa - fissuras'),('estrutura|concreto','Estrutura principal - solidez e segurança')]
    for pat,cat in rules:
        if re.search(pat,t): return cat
    return 'Outros / análise técnica'

def warranty_analysis(emp,cat,de,da):
    row=Garantia.query.filter(Garantia.item==cat,Garantia.empreendimento.in_([emp,'Todos'])).order_by(case((Garantia.empreendimento==emp,0),else_=1)).first()
    if not row: return {'resultado':'Necessita vistoria','prazo':'','fundamentacao':'Item não localizado na base resumida. Consulte o manual completo e realize vistoria técnica.'}
    elapsed=None
    try:
        d1=datetime.strptime(de,'%Y-%m-%d').date(); d2=datetime.strptime(da,'%Y-%m-%d').date(); elapsed=(d2.year-d1.year)*12+d2.month-d1.month-(1 if d2.day<d1.day else 0)
    except: pass
    if row.prazo_meses==0: result='Improcedente'
    elif elapsed is None: result='Necessita vistoria'
    elif elapsed<=row.prazo_meses: result=row.classificacao_padrao or 'Necessita vistoria'
    else: result='Fora do prazo / analisar improcedência'
    return {'resultado':result,'prazo':f'{row.prazo_meses} meses' if row.prazo_meses else 'No ato da entrega','fundamentacao':row.descricao,'fonte':row.fonte,'meses_decorridos':elapsed}

@app.context_processor
def context(): return dict(can_edit=can_edit(),is_admin=is_admin(),perfil=session.get('perfil'),now=datetime.now())

@app.route('/login',methods=['GET','POST'])
def login():
    if request.method=='POST':
        u=Usuario.query.filter(func.lower(Usuario.email)==request.form.get('email','').strip().lower(),Usuario.ativo==True).first()
        if u and check(u.senha,request.form.get('senha','')):
            session.clear(); session.update(user_id=u.id,nome=u.nome,perfil=u.perfil); return redirect(url_for('dashboard'))
        flash('E-mail ou senha inválidos.','erro')
    return render_template('login.html')
@app.route('/sair')
def logout(): session.clear(); return redirect(url_for('login'))

@app.route('/')
@login_required
def dashboard():
    total=Chamado.query.count(); final=Chamado.query.filter(func.lower(Chamado.status).like('%final%')).count(); impro=Chamado.query.filter(func.lower(Chamado.classificacao).like('%improced%')).count(); abertos=total-final
    por_emp=db.session.query(Chamado.empreendimento,func.count(Chamado.id)).group_by(Chamado.empreendimento).order_by(func.count(Chamado.id).desc()).all()
    recentes=Chamado.query.order_by(Chamado.id.desc()).limit(8).all()
    fin=None
    if is_admin():
        entradas=db.session.query(func.coalesce(func.sum(Financeiro.valor),0)).filter(Financeiro.tipo=='Entrada',Financeiro.status=='Pago/Recebido').scalar(); saidas=db.session.query(func.coalesce(func.sum(Financeiro.valor),0)).filter(Financeiro.tipo=='Saída',Financeiro.status=='Pago/Recebido').scalar(); fin={'entradas':entradas,'saidas':saidas}
    em_campo=Atendimento.query.filter_by(tipo='checkin').count()-Atendimento.query.filter_by(tipo='checkout').count()
    return render_template('dashboard.html',k={'total':total,'finalizados':final,'abertos':abertos,'improcedentes':impro,'em_campo':max(em_campo,0)},por_emp=por_emp,recentes=recentes,fin=fin)

@app.route('/chamados')
@login_required
def chamados():
    q=request.args.get('q','').strip(); emp=request.args.get('empreendimento',''); status=request.args.get('status',''); qry=Chamado.query
    if q: qry=qry.filter(or_(Chamado.sat.ilike(f'%{q}%'),Chamado.solicitante.ilike(f'%{q}%'),Chamado.unidade.ilike(f'%{q}%'),Chamado.problema.ilike(f'%{q}%')))
    if emp: qry=qry.filter_by(empreendimento=emp)
    if status: qry=qry.filter_by(status=status)
    return render_template('chamados.html',rows=qry.order_by(Chamado.id.desc()).all(),emps=Empreendimento.query.order_by(Empreendimento.nome).all(),statuses=[x[0] for x in db.session.query(Chamado.status).distinct().order_by(Chamado.status)])

@app.route('/chamados/novo',methods=['GET','POST'])
@login_required
@allow('Administrador','Líder','Equipe')
def novo_chamado():
    if request.method=='POST':
        cat=request.form.get('categoria') or guess_category(request.form.get('problema')); da=request.form.get('data_recebido') or date.today().isoformat(); ana=warranty_analysis(request.form.get('empreendimento'),cat,request.form.get('data_entrega'),da)
        c=Chamado(sat=request.form.get('sat'),empreendimento=request.form.get('empreendimento'),data_recebido=da,solicitante=request.form.get('solicitante'),unidade=request.form.get('unidade'),problema=request.form.get('problema'),contato=request.form.get('contato'),status=request.form.get('status') or 'Aberto',classificacao=request.form.get('classificacao') or ana['resultado'],categoria=cat,data_entrega=request.form.get('data_entrega'),analise_garantia=ana['resultado'],fundamentacao=ana['fundamentacao'],responsavel=request.form.get('responsavel'),observacoes=request.form.get('observacoes'))
        db.session.add(c); db.session.commit(); log(c.id,'Chamado cadastrado'); flash('Chamado cadastrado com sucesso.','ok'); return redirect(url_for('ver_chamado',id=c.id))
    return render_template('chamado_form.html',emps=Empreendimento.query.order_by(Empreendimento.nome).all(),cats=[x[0] for x in db.session.query(Garantia.item).distinct().order_by(Garantia.item)])


@app.route('/chamados/importar',methods=['GET','POST'])
@login_required
@allow('Administrador','Líder','Equipe')
def importar_chamados():
    if request.method=='POST':
        f=request.files.get('planilha')
        if not f or not f.filename:
            flash('Selecione uma planilha.','erro'); return redirect(url_for('importar_chamados'))
        try:
            rows,mapping=read_import_file(f)
            sats=[r.get('sat','').strip() for r in rows if r.get('sat')]
            existing={x[0] for x in db.session.query(Chamado.sat).filter(Chamado.sat.in_(sats)).all() if x[0]} if sats else set()
            seen=set()
            for r in rows:
                sat=r.get('sat','').strip()
                if sat and (sat in existing or sat in seen): r['_duplicada']=True
                else: r['_duplicada']=False
                if sat: seen.add(sat)
            token=uuid.uuid4().hex
            (IMPORT_DIR/f'{token}.json').write_text(json.dumps(rows,ensure_ascii=False),encoding='utf-8')
            return render_template('importar_chamados.html',rows=rows,token=token,mapping=mapping,preview=True)
        except Exception as e:
            flash(str(e),'erro')
    return render_template('importar_chamados.html',preview=False)

@app.post('/chamados/importar/confirmar')
@login_required
@allow('Administrador','Líder','Equipe')
def confirmar_importacao():
    token=re.sub(r'[^a-f0-9]','',request.form.get('token',''))
    path=IMPORT_DIR/f'{token}.json'
    if not path.exists():
        flash('A prévia expirou. Envie a planilha novamente.','erro'); return redirect(url_for('importar_chamados'))
    rows=json.loads(path.read_text(encoding='utf-8')); imported=duplicates=errors=0
    for r in rows:
        if r.get('_erro'): errors+=1; continue
        sat=(r.get('sat') or '').strip()
        if sat and Chamado.query.filter_by(sat=sat).first(): duplicates+=1; continue
        emp=(r.get('empreendimento') or '').strip()
        if emp and not Empreendimento.query.filter_by(nome=emp).first(): db.session.add(Empreendimento(nome=emp))
        cat=r.get('categoria') or guess_category(r.get('problema'))
        da=r.get('data_recebido') or date.today().isoformat()
        ana=warranty_analysis(emp,cat,r.get('data_entrega'),da)
        c=Chamado(sat=sat,empreendimento=emp,data_recebido=da,solicitante=r.get('solicitante'),unidade=r.get('unidade'),problema=r.get('problema'),contato=r.get('contato'),status=r.get('status') or 'Aberto',classificacao=r.get('classificacao') or ana['resultado'],categoria=cat,data_entrega=r.get('data_entrega'),analise_garantia=ana['resultado'],fundamentacao=ana['fundamentacao'],responsavel=r.get('responsavel'),observacoes=r.get('observacoes'))
        db.session.add(c); db.session.flush(); db.session.add(Historico(chamado_id=c.id,usuario=session.get('nome','Sistema'),acao='Chamado importado por planilha'))
        imported+=1
    db.session.commit()
    try: path.unlink()
    except OSError: pass
    flash(f'Importação concluída: {imported} chamados incluídos, {duplicates} duplicados ignorados e {errors} linhas com erro.','ok')
    return redirect(url_for('chamados'))

@app.route('/chamados/modelo-importacao.xlsx')
@login_required
def modelo_importacao():
    wb=Workbook(); ws=wb.active; ws.title='Chamados'
    headers=['SAT','Empreendimento','Data de abertura','Morador','Unidade','Problema','Contato','Status','Classificação','Categoria','Data de entrega','Responsável','Observações']
    ws.append(headers); ws.append(['26576','Vittace Reserva','20/07/2026','Maria da Silva','Torre 06 - Apto 403','Infiltração no teto do banheiro','(43) 99999-9999','Aberto','','','15/03/2023','',''])
    ws.freeze_panes='A2'; ws.auto_filter.ref=f'A1:M2'
    widths=[14,24,18,24,22,48,20,16,20,28,18,22,38]
    for i,w in enumerate(widths,1): ws.column_dimensions[chr(64+i)].width=w
    out=io.BytesIO(); wb.save(out); out.seek(0)
    return send_file(out,as_attachment=True,download_name='Modelo_Importacao_BR_SmartSAT.xlsx',mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/chamados/<int:id>',methods=['GET','POST'])
@login_required
def ver_chamado(id):
    c=Chamado.query.get_or_404(id)
    if request.method=='POST':
        if not can_edit(): abort(403)
        antes=snapshot(c,['sat','empreendimento','data_recebido','solicitante','unidade','problema','contato','status','classificacao','categoria','data_entrega','responsavel','observacoes'])
        for field in ['sat','empreendimento','data_recebido','solicitante','unidade','problema','contato','status','classificacao','categoria','data_entrega','responsavel','observacoes']:
            if field in request.form: setattr(c,field,request.form.get(field))
        db.session.commit(); log(id,f"Chamado atualizado para status: {c.status}"); audit('Chamado',id,'EDITADO',antes+' => '+snapshot(c,['sat','empreendimento','data_recebido','solicitante','unidade','problema','contato','status','classificacao','categoria','data_entrega','responsavel','observacoes'])); flash('Chamado atualizado.','ok'); return redirect(url_for('ver_chamado',id=id))
    return render_template('chamado_detalhe.html',c=c,fotos=Foto.query.filter_by(chamado_id=id).order_by(Foto.id.desc()).all(),assinaturas=Assinatura.query.filter_by(chamado_id=id).order_by(Assinatura.id.desc()).all(),mats=Material.query.filter_by(chamado_id=id).order_by(Material.id.desc()).all(),hist=Historico.query.filter_by(chamado_id=id).order_by(Historico.id.desc()).limit(50).all(),atendimentos=Atendimento.query.filter_by(chamado_id=id).order_by(Atendimento.id.desc()).all(),relatorio=RelatorioTecnico.query.filter_by(chamado_id=id).first(),cats=[x[0] for x in db.session.query(Garantia.item).distinct().order_by(Garantia.item)])


@app.post('/chamados/<int:id>/excluir')
@login_required
@allow('Administrador','Líder')
def excluir_chamado(id):
    c=Chamado.query.get_or_404(id)
    detalhes=snapshot(c,['sat','empreendimento','solicitante','unidade','status'])
    # Remove arquivos físicos e registros vinculados.
    for f in Foto.query.filter_by(chamado_id=id).all():
        try: (UPLOAD/f.arquivo).unlink(missing_ok=True)
        except Exception: pass
        db.session.delete(f)
    for a in Assinatura.query.filter_by(chamado_id=id).all():
        try: (UPLOAD/a.arquivo).unlink(missing_ok=True)
        except Exception: pass
        db.session.delete(a)
    Material.query.filter_by(chamado_id=id).delete(synchronize_session=False)
    Atendimento.query.filter_by(chamado_id=id).delete(synchronize_session=False)
    RelatorioTecnico.query.filter_by(chamado_id=id).delete(synchronize_session=False)
    Historico.query.filter_by(chamado_id=id).delete(synchronize_session=False)
    db.session.delete(c); db.session.commit(); audit('Chamado',id,'EXCLUÍDO',detalhes)
    flash('Chamado excluído com sucesso.','ok'); return redirect(url_for('chamados'))

@app.post('/fotos/<int:id>/editar')
@login_required
@allow('Administrador','Líder','Equipe','Técnico')
def editar_foto(id):
    f=Foto.query.get_or_404(id); antes=f.descricao or ''
    f.descricao=request.form.get('descricao') or f.descricao; db.session.commit(); log(f.chamado_id,f'Foto editada: {antes} para {f.descricao}'); audit('Foto',id,'EDITADA',antes+' => '+(f.descricao or ''))
    flash('Descrição da foto atualizada.','ok'); return redirect(url_for('ver_chamado',id=f.chamado_id))

@app.post('/fotos/<int:id>/excluir')
@login_required
@allow('Administrador','Líder','Equipe','Técnico')
def excluir_foto(id):
    f=Foto.query.get_or_404(id); cid=f.chamado_id; detalhes=f'{f.descricao} | {f.arquivo}'
    try: (UPLOAD/f.arquivo).unlink(missing_ok=True)
    except Exception: pass
    db.session.delete(f); db.session.commit(); log(cid,'Foto excluída'); audit('Foto',id,'EXCLUÍDA',detalhes)
    flash('Foto excluída.','ok'); return redirect(url_for('ver_chamado',id=cid))

@app.post('/assinaturas/<int:id>/editar')
@login_required
@allow('Administrador','Líder','Equipe','Técnico')
def editar_assinatura(id):
    a=Assinatura.query.get_or_404(id); antes=snapshot(a,['nome_cliente','documento','observacao'])
    a.nome_cliente=request.form.get('nome_cliente') or a.nome_cliente; a.documento=request.form.get('documento'); a.observacao=request.form.get('observacao'); db.session.commit(); log(a.chamado_id,'Dados da assinatura atualizados'); audit('Assinatura',id,'EDITADA',antes+' => '+snapshot(a,['nome_cliente','documento','observacao']))
    flash('Dados da assinatura atualizados.','ok'); return redirect(url_for('ver_chamado',id=a.chamado_id))

@app.post('/assinaturas/<int:id>/excluir')
@login_required
@allow('Administrador','Líder','Equipe','Técnico')
def excluir_assinatura(id):
    a=Assinatura.query.get_or_404(id); cid=a.chamado_id; detalhes=snapshot(a,['nome_cliente','documento','assinado_em'])
    try: (UPLOAD/a.arquivo).unlink(missing_ok=True)
    except Exception: pass
    db.session.delete(a); db.session.commit(); log(cid,'Assinatura excluída para nova coleta'); audit('Assinatura',id,'EXCLUÍDA',detalhes)
    flash('Assinatura excluída. Uma nova assinatura já pode ser coletada.','ok'); return redirect(url_for('ver_chamado',id=cid))

@app.post('/materiais/<int:id>/editar')
@login_required
@allow('Administrador','Líder','Equipe','Técnico')
def editar_material(id):
    m=Material.query.get_or_404(id); antes=snapshot(m,['item','quantidade','status','observacoes'])
    for field in ['item','quantidade','status','observacoes']: setattr(m,field,request.form.get(field))
    db.session.commit(); log(m.chamado_id,'Material editado'); audit('Material',id,'EDITADO',antes+' => '+snapshot(m,['item','quantidade','status','observacoes']))
    flash('Material atualizado.','ok'); return redirect(url_for('ver_chamado',id=m.chamado_id))

@app.post('/materiais/<int:id>/excluir')
@login_required
@allow('Administrador','Líder','Equipe','Técnico')
def excluir_material(id):
    m=Material.query.get_or_404(id); cid=m.chamado_id; detalhes=snapshot(m,['item','quantidade','status'])
    db.session.delete(m); db.session.commit(); log(cid,'Material excluído'); audit('Material',id,'EXCLUÍDO',detalhes)
    flash('Material excluído.','ok'); return redirect(url_for('ver_chamado',id=cid))

@app.post('/chamados/<int:id>/analisar')
@login_required
@allow('Administrador','Líder','Equipe')
def analisar_chamado(id):
    c=Chamado.query.get_or_404(id); cat=request.form.get('categoria') or guess_category(c.problema); a=warranty_analysis(c.empreendimento,cat,c.data_entrega,c.data_recebido); c.categoria=cat;c.analise_garantia=a['resultado'];c.fundamentacao=a['fundamentacao']+((' Fonte: '+a.get('fonte','')) if a.get('fonte') else '');db.session.commit();log(id,'Garantia analisada');return redirect(url_for('ver_chamado',id=id))
@app.post('/chamados/<int:id>/foto')
@login_required
@allow('Administrador','Líder','Equipe','Técnico')
def foto(id):
    Chamado.query.get_or_404(id)
    f=request.files.get('foto')
    if not f or not f.filename:
        flash('Selecione ou tire uma foto.','erro'); return redirect(url_for('ver_chamado',id=id))
    loc=parse_location(request.form)
    if not loc:
        flash('A localização é obrigatória. Ative o GPS e permita o acesso à localização.','erro'); return redirect(url_for('ver_chamado',id=id))
    lat,lon,acc=loc
    name=f'{uuid.uuid4().hex}_{secure_filename(f.filename)}'; f.save(UPLOAD/name)
    db.session.add(Foto(chamado_id=id,arquivo=name,descricao=request.form.get('descricao'),latitude=lat,longitude=lon,precisao=acc,capturado_em=parse_client_time(request.form.get('capturado_em')),usuario=session.get('nome'),dispositivo=(request.form.get('dispositivo') or '')[:255]))
    db.session.commit();log(id,f'Foto adicionada com GPS: {lat:.6f}, {lon:.6f}');flash('Foto e localização registradas.','ok');return redirect(url_for('ver_chamado',id=id))

@app.post('/chamados/<int:id>/assinatura')
@login_required
@allow('Administrador','Líder','Equipe','Técnico')
def assinatura(id):
    Chamado.query.get_or_404(id)
    data=request.form.get('assinatura_data','')
    if not data.startswith('data:image/png;base64,'):
        flash('Faça a assinatura no quadro antes de salvar.','erro'); return redirect(url_for('ver_chamado',id=id))
    try:
        raw=base64.b64decode(data.split(',',1)[1])
        name=f'assinatura_{id}_{uuid.uuid4().hex}.png'
        (UPLOAD/name).write_bytes(raw)
        loc=parse_location(request.form)
        if not loc:
            flash('A localização é obrigatória para registrar a assinatura. Ative o GPS e tente novamente.','erro'); return redirect(url_for('ver_chamado',id=id))
        lat,lon,acc=loc
        db.session.add(Assinatura(chamado_id=id,arquivo=name,nome_cliente=request.form.get('nome_cliente'),documento=request.form.get('documento'),observacao=request.form.get('observacao'),latitude=lat,longitude=lon,precisao=acc,assinado_em=parse_client_time(request.form.get('assinado_em')),usuario=session.get('nome'),dispositivo=(request.form.get('dispositivo') or '')[:255]))
        db.session.commit(); log(id,f"Assinatura do cliente registrada com GPS: {lat:.6f}, {lon:.6f}")
        flash('Assinatura registrada com sucesso.','ok')
    except Exception:
        flash('Não foi possível salvar a assinatura. Tente novamente.','erro')
    return redirect(url_for('ver_chamado',id=id))

@app.post('/chamados/<int:id>/material')
@login_required
@allow('Administrador','Líder','Equipe','Técnico')
def material(id):
    db.session.add(Material(chamado_id=id,item=request.form.get('item'),quantidade=request.form.get('quantidade'),status=request.form.get('status'),observacoes=request.form.get('observacoes')));db.session.commit();log(id,'Material registrado');return redirect(url_for('ver_chamado',id=id))


@app.post('/chamados/<int:id>/atendimento/<tipo>')
@login_required
@allow('Administrador','Líder','Equipe','Técnico')
def registrar_atendimento(id,tipo):
    Chamado.query.get_or_404(id)
    if tipo not in ('checkin','checkout'): abort(400)
    loc=parse_location(request.form)
    if not loc:
        flash('A localização é obrigatória para registrar chegada ou saída.','erro'); return redirect(url_for('ver_chamado',id=id))
    lat,lon,acc=loc
    db.session.add(Atendimento(chamado_id=id,tipo=tipo,latitude=lat,longitude=lon,precisao=acc,usuario=session.get('nome'),dispositivo=(request.form.get('dispositivo') or '')[:255],observacao=request.form.get('observacao')))
    db.session.commit(); log(id,'Check-in registrado' if tipo=='checkin' else 'Check-out registrado')
    flash(('Chegada' if tipo=='checkin' else 'Saída')+' registrada com GPS.','ok'); return redirect(url_for('ver_chamado',id=id))

@app.post('/chamados/<int:id>/relatorio')
@login_required
@allow('Administrador','Líder','Equipe','Técnico')
def salvar_relatorio(id):
    Chamado.query.get_or_404(id)
    r=RelatorioTecnico.query.filter_by(chamado_id=id).first() or RelatorioTecnico(chamado_id=id)
    r.diagnostico=request.form.get('diagnostico'); r.servico_executado=request.form.get('servico_executado'); r.testes_realizados=request.form.get('testes_realizados'); r.conclusao=request.form.get('conclusao'); r.usuario=session.get('nome')
    db.session.add(r); db.session.commit(); log(id,'Relatório técnico atualizado'); flash('Relatório técnico salvo.','ok'); return redirect(url_for('ver_chamado',id=id))

@app.route('/chamados/<int:id>/pdf')
@login_required
def pdf_chamado(id):
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
    from reportlab.lib.utils import ImageReader
    c=Chamado.query.get_or_404(id); rel=RelatorioTecnico.query.filter_by(chamado_id=id).first(); ass=Assinatura.query.filter_by(chamado_id=id).order_by(Assinatura.id.desc()).first(); fotos=Foto.query.filter_by(chamado_id=id).order_by(Foto.id).all()
    buf=io.BytesIO(); pdf=canvas.Canvas(buf,pagesize=A4); W,H=A4; y=H-45
    def line(txt,bold=False,size=10):
        nonlocal y
        pdf.setFont('Helvetica-Bold' if bold else 'Helvetica',size)
        for part in [txt[i:i+95] for i in range(0,len(txt or ''),95)] or ['']:
            if y<60: pdf.showPage(); y=H-45
            pdf.drawString(40,y,part); y-=14
    line('BR SMARTSAT ENTERPRISE 12.0 - RELATORIO DE ATENDIMENTO',True,14); line(f'SAT: {c.sat or c.id}',True); line(f'Empreendimento: {c.empreendimento}'); line(f'Unidade: {c.unidade}'); line(f'Morador: {c.solicitante}'); line(f'Status: {c.status} | Classificacao: {c.classificacao}'); line('Problema informado:',True); line(c.problema or '-')
    if rel:
        line('Diagnostico tecnico:',True); line(rel.diagnostico or '-'); line('Servico executado:',True); line(rel.servico_executado or '-'); line('Testes realizados:',True); line(rel.testes_realizados or '-'); line('Conclusao:',True); line(rel.conclusao or '-')
    line(f'Fotos registradas: {len(fotos)}',True)
    for f in fotos[:6]:
        path=UPLOAD/f.arquivo
        if path.exists():
            if y<180: pdf.showPage(); y=H-45
            try: pdf.drawImage(ImageReader(str(path)),40,y-130,width=180,height=125,preserveAspectRatio=True,anchor='sw'); pdf.drawString(230,y-20,(f.descricao or 'Foto')[:50]); pdf.drawString(230,y-38,f'GPS: {f.latitude}, {f.longitude}'); y-=145
            except: pass
    if ass:
        line('Assinatura do cliente:',True); line(f'{ass.nome_cliente or "Cliente"} - {ass.documento or "documento nao informado"}')
        path=UPLOAD/ass.arquivo
        if path.exists():
            if y<160: pdf.showPage(); y=H-45
            try: pdf.drawImage(ImageReader(str(path)),40,y-110,width=300,height=100,preserveAspectRatio=True,anchor='sw'); y-=120
            except: pass
    line(f'Codigo de validacao: SAT-{c.id}-{(c.sat or c.id)}',True); pdf.save(); buf.seek(0)
    return send_file(buf,as_attachment=True,download_name=f'SAT_{c.sat or c.id}_BR_SmartSAT.pdf',mimetype='application/pdf')

@app.route('/validar/<int:id>')
def validar(id):
    c=Chamado.query.get_or_404(id); ass=Assinatura.query.filter_by(chamado_id=id).order_by(Assinatura.id.desc()).first(); return render_template('validar.html',c=c,ass=ass)

@app.route('/manifest.webmanifest')
def manifest(): return send_file(BASE/'static'/'manifest.webmanifest',mimetype='application/manifest+json')
@app.route('/service-worker.js')
def sw(): return send_file(BASE/'static'/'service-worker.js',mimetype='application/javascript')

@app.route('/garantias')
@login_required
def garantias():
    emp=request.args.get('empreendimento','');q=request.args.get('q','');qry=Garantia.query
    if emp: qry=qry.filter(Garantia.empreendimento.in_([emp,'Todos']))
    if q: qry=qry.filter(or_(Garantia.item.ilike(f'%{q}%'),Garantia.descricao.ilike(f'%{q}%')))
    return render_template('garantias.html',rows=qry.order_by(Garantia.empreendimento,Garantia.item,Garantia.prazo_meses).all())

@app.route('/financeiro',methods=['GET','POST'])
@login_required
@allow('Administrador')
def financeiro():
    if request.method=='POST':
        r=Financeiro(data=request.form.get('data'),tipo=request.form.get('tipo'),descricao=request.form.get('descricao'),empreendimento=request.form.get('empreendimento'),categoria=request.form.get('categoria'),valor=float(request.form.get('valor') or 0),status=request.form.get('status'),forma_pagamento=request.form.get('forma_pagamento')); db.session.add(r);db.session.commit();audit('Financeiro',r.id,'CRIADO',snapshot(r,['data','tipo','descricao','empreendimento','categoria','valor','status','forma_pagamento']));flash('Lançamento financeiro salvo.','ok');return redirect(url_for('financeiro'))
    rows=Financeiro.query.order_by(Financeiro.data.desc(),Financeiro.id.desc()).all();ent=db.session.query(func.coalesce(func.sum(Financeiro.valor),0)).filter_by(tipo='Entrada',status='Pago/Recebido').scalar();sai=db.session.query(func.coalesce(func.sum(Financeiro.valor),0)).filter_by(tipo='Saída',status='Pago/Recebido').scalar();return render_template('financeiro.html',rows=rows,sums={'entradas':ent,'saidas':sai},emps=Empreendimento.query.order_by(Empreendimento.nome).all())

@app.route('/financeiro/<int:id>/editar',methods=['GET','POST'])
@login_required
@allow('Administrador')
def editar_financeiro(id):
    r=Financeiro.query.get_or_404(id)
    if request.method=='POST':
        antes=snapshot(r,['data','tipo','descricao','empreendimento','categoria','valor','status','forma_pagamento'])
        r.data=request.form.get('data'); r.tipo=request.form.get('tipo'); r.descricao=request.form.get('descricao'); r.empreendimento=request.form.get('empreendimento'); r.categoria=request.form.get('categoria'); r.valor=float(request.form.get('valor') or 0); r.status=request.form.get('status'); r.forma_pagamento=request.form.get('forma_pagamento'); db.session.commit(); audit('Financeiro',id,'EDITADO',antes+' => '+snapshot(r,['data','tipo','descricao','empreendimento','categoria','valor','status','forma_pagamento'])); flash('Lançamento atualizado.','ok'); return redirect(url_for('financeiro'))
    return render_template('financeiro_editar.html',r=r,emps=Empreendimento.query.order_by(Empreendimento.nome).all())

@app.post('/financeiro/<int:id>/excluir')
@login_required
@allow('Administrador')
def excluir_financeiro(id):
    r=Financeiro.query.get_or_404(id); detalhes=snapshot(r,['data','tipo','descricao','empreendimento','categoria','valor','status','forma_pagamento']); db.session.delete(r); db.session.commit(); audit('Financeiro',id,'EXCLUÍDO',detalhes); flash('Lançamento excluído e saldo recalculado.','ok'); return redirect(url_for('financeiro'))

@app.route('/empreendimentos',methods=['GET','POST'])
@login_required
@allow('Administrador','Líder')
def empreendimentos():
    if request.method=='POST':
        if not Empreendimento.query.filter_by(nome=request.form.get('nome')).first(): db.session.add(Empreendimento(nome=request.form.get('nome'),data_entrega=request.form.get('data_entrega'),observacoes=request.form.get('observacoes')));db.session.commit()
        return redirect(url_for('empreendimentos'))
    return render_template('empreendimentos.html',rows=Empreendimento.query.order_by(Empreendimento.nome).all())

@app.route('/usuarios',methods=['GET','POST'])
@login_required
@allow('Administrador')
def usuarios():
    if request.method=='POST':
        email=request.form.get('email','').lower().strip()
        if Usuario.query.filter_by(email=email).first(): flash('E-mail já cadastrado.','erro')
        else: db.session.add(Usuario(nome=request.form.get('nome'),email=email,senha=ph(request.form.get('senha')),perfil=request.form.get('perfil'),ativo=True));db.session.commit();flash('Usuário cadastrado.','ok')
        return redirect(url_for('usuarios'))
    return render_template('usuarios.html',rows=Usuario.query.order_by(Usuario.nome).all())
@app.post('/usuarios/<int:id>/alternar')
@login_required
@allow('Administrador')
def alternar_usuario(id):
    u=Usuario.query.get_or_404(id)
    if u.id==session.get('user_id'): flash('Você não pode desativar seu próprio usuário.','erro')
    else: u.ativo=not u.ativo;db.session.commit()
    return redirect(url_for('usuarios'))

@app.route('/exportar')
@login_required
def exportar():
    empreendimento_filtro=(request.args.get('empreendimento') or '').strip()
    qry=Chamado.query
    if empreendimento_filtro:
        qry=qry.filter(Chamado.empreendimento==empreendimento_filtro)
    rows=qry.order_by(Chamado.empreendimento,Chamado.status,Chamado.data_recebido,Chamado.id).all()

    wb=Workbook()
    resumo=wb.active
    resumo.title='Resumo Geral'

    azul='0B3B5A'; azul_claro='D9EAF4'; branco='FFFFFF'; cinza='E5E7EB'
    borda=Border(left=Side(style='thin',color='D1D5DB'),right=Side(style='thin',color='D1D5DB'),top=Side(style='thin',color='D1D5DB'),bottom=Side(style='thin',color='D1D5DB'))
    status_cores={
      'aberto':'BDD7EE','aberta':'BDD7EE','em execucao':'FFF2CC','em execução':'FFF2CC','em andamento':'FFF2CC',
      'aguardando material':'FCE4D6','aguardando morador':'E4DFEC','aguardando cliente':'E4DFEC',
      'finalizado':'C6E0B4','finalizada':'C6E0B4','concluido':'C6E0B4','concluído':'C6E0B4',
      'improcedente':'F4CCCC','procedente':'D9EAD3','em analise':'D9D9D9','em análise':'D9D9D9','cancelado':'F4CCCC','cancelada':'F4CCCC'
    }
    def safe_sheet_name(name,used):
        base=re.sub(r'[\/*?:\[\]]','-',(name or 'Sem empreendimento')).strip()[:31] or 'Sem empreendimento'
        candidate=base; n=2
        while candidate in used:
            suffix=f' ({n})'; candidate=(base[:31-len(suffix)]+suffix); n+=1
        used.add(candidate); return candidate
    def style_header(ws,row=1):
        for cell in ws[row]:
            cell.fill=PatternFill('solid',fgColor=azul); cell.font=Font(color=branco,bold=True); cell.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); cell.border=borda
        ws.row_dimensions[row].height=28
    def format_status(cell):
        key=norm_header(cell.value)
        color=status_cores.get(key)
        if color: cell.fill=PatternFill('solid',fgColor=color)
        cell.font=Font(bold=True); cell.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); cell.border=borda

    # Resumo geral
    resumo.merge_cells('A1:F1'); resumo['A1']='BR SmartSAT — Resumo de Chamados'
    resumo['A1'].fill=PatternFill('solid',fgColor=azul); resumo['A1'].font=Font(color=branco,bold=True,size=16); resumo['A1'].alignment=Alignment(horizontal='center'); resumo.row_dimensions[1].height=30
    resumo.append([])
    resumo.append(['Indicador','Quantidade'])
    total=len(rows)
    resumo.append(['Total de SATs',total])
    status_counts={}
    emp_counts={}
    for r in rows:
        st=(r.status or 'Sem status').strip(); status_counts[st]=status_counts.get(st,0)+1
        emp=(r.empreendimento or 'Sem empreendimento').strip(); emp_counts[emp]=emp_counts.get(emp,0)+1
    for st,count in sorted(status_counts.items(),key=lambda x:x[0].lower()): resumo.append([st,count])
    style_header(resumo,3)
    for row in resumo.iter_rows(min_row=4,max_row=3+1+len(status_counts),min_col=1,max_col=2):
        for c in row: c.border=borda
        if row[0].value!='Total de SATs': format_status(row[0])
    start=6+len(status_counts)
    resumo.cell(start,1,'Empreendimento'); resumo.cell(start,2,'Quantidade')
    for c in resumo[start]: c.fill=PatternFill('solid',fgColor=azul); c.font=Font(color=branco,bold=True); c.border=borda
    for i,(emp,count) in enumerate(sorted(emp_counts.items(),key=lambda x:x[0].lower()),start=start+1):
        resumo.cell(i,1,emp); resumo.cell(i,2,count); resumo.cell(i,1).border=borda; resumo.cell(i,2).border=borda
    resumo.column_dimensions['A'].width=34; resumo.column_dimensions['B'].width=16; resumo.freeze_panes='A4'

    headers=['SAT','Data de abertura','Empreendimento','Unidade','Morador / Solicitante','Contato','Problema / Solicitação','Categoria','Classificação','Status','Responsável','Observações','Última atualização']
    grouped={}
    for r in rows: grouped.setdefault((r.empreendimento or 'Sem empreendimento').strip(),[]).append(r)
    used={'Resumo Geral'}
    for emp,items in sorted(grouped.items(),key=lambda x:x[0].lower()):
        ws=wb.create_sheet(safe_sheet_name(emp,used))
        ws.append(headers); style_header(ws)
        for r in items:
            ws.append([
              r.sat or r.id, excel_date(r.data_recebido), r.empreendimento or '', r.unidade or '', r.solicitante or '', r.contato or '',
              r.problema or '', r.categoria or '', r.classificacao or '', r.status or 'Sem status', r.responsavel or '', r.observacoes or '',
              r.atualizado_em.strftime('%d/%m/%Y %H:%M') if r.atualizado_em else ''
            ])
            rr=ws.max_row
            for c in ws[rr]: c.border=borda; c.alignment=Alignment(vertical='top',wrap_text=True)
            format_status(ws.cell(rr,10))
        ws.auto_filter.ref=f'A1:M{ws.max_row}'
        ws.freeze_panes='A2'
        widths=[13,16,28,16,26,18,48,34,18,20,24,40,20]
        for idx,width in enumerate(widths,1): ws.column_dimensions[get_column_letter(idx)].width=width
        for rr in range(2,ws.max_row+1): ws.row_dimensions[rr].height=42
        ws.sheet_view.showGridLines=False

    out=io.BytesIO(); wb.save(out); out.seek(0)
    nome='Chamados_BR_SmartSAT_por_empreendimento.xlsx' if not empreendimento_filtro else f'Chamados_{secure_filename(empreendimento_filtro)}.xlsx'
    return send_file(out,as_attachment=True,download_name=nome,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/api/analisar',methods=['POST'])
@login_required
def api_analisar():
    d=request.get_json(force=True);cat=d.get('categoria') or guess_category(d.get('problema'));return jsonify({'categoria':cat,**warranty_analysis(d.get('empreendimento'),cat,d.get('data_entrega'),d.get('data_abertura'))})

with app.app_context(): init_db()
if __name__=='__main__': app.run(host='0.0.0.0',port=int(os.environ.get('PORT',5000)),debug=False)
