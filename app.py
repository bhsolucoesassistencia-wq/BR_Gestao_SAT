
import os, sqlite3, csv, io, base64, uuid
from datetime import datetime, date
from functools import wraps
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file, jsonify
import hashlib

def generate_password_hash(s): return hashlib.sha256(s.encode('utf-8')).hexdigest()
def check_password_hash(h,s): return h==hashlib.sha256(s.encode('utf-8')).hexdigest()
from werkzeug.utils import secure_filename

BASE=os.path.dirname(os.path.abspath(__file__))
DB=os.path.join(BASE,"data","gestao.db")
UPLOAD=os.path.join(BASE,"static","uploads")
os.makedirs(UPLOAD,exist_ok=True)

app=Flask(__name__)
app.secret_key=os.environ.get("SECRET_KEY","br-gestao-sat-8-chave-local")
app.config["MAX_CONTENT_LENGTH"]=30*1024*1024

def db():
    con=sqlite3.connect(DB)
    con.row_factory=sqlite3.Row
    return con

def login_required(fn):
    @wraps(fn)
    def w(*a,**k):
        if not session.get("user_id"): return redirect(url_for("login"))
        return fn(*a,**k)
    return w

def admin_required(fn):
    @wraps(fn)
    def w(*a,**k):
        if session.get("perfil")!="Administrador":
            flash("Acesso restrito ao administrador.","erro")
            return redirect(url_for("dashboard"))
        return fn(*a,**k)
    return w

def normalize_status(s):
    s=(s or "").strip()
    return s or "Aberto"

def guess_category(text):
    t=(text or "").lower()
    rules=[
      ("infiltra|vazamento|umidade|mofo","Instalações hidráulicas - vedação / vazamento"),
      ("janela|esquadria|vedação","Esquadrias de alumínio - vedação e funcionamento"),
      ("telha|telhado|cobertura|calha","Cobertura e telhados"),
      ("drywall|gesso","Drywall - fissuras"),
      ("pintura|descasc|empol|tinta","Pintura interna"),
      ("fissura|trinca|rachadura","Revestimento em argamassa - fissuras"),
      ("estrutura|concreto","Estrutura principal - solidez e segurança"),
    ]
    import re
    for pat,cat in rules:
        if re.search(pat,t): return cat
    return "Outros / análise técnica"

def warranty_analysis(empreendimento,categoria,data_entrega,data_abertura):
    con=db()
    row=con.execute("""SELECT * FROM garantias WHERE item=? AND empreendimento IN (?, 'Todos')
                       ORDER BY CASE WHEN empreendimento=? THEN 0 ELSE 1 END LIMIT 1""",
                    (categoria,empreendimento,empreendimento)).fetchone()
    if not row:
        return {"resultado":"Necessita vistoria","prazo":"","fundamentacao":"Item não localizado na base resumida. Consulte o manual completo e faça vistoria técnica.","regra":None}
    months=row["prazo_meses"]
    elapsed=None
    try:
        d1=datetime.strptime(data_entrega,"%Y-%m-%d").date()
        d2=datetime.strptime(data_abertura,"%Y-%m-%d").date()
        elapsed=(d2.year-d1.year)*12+d2.month-d1.month-(1 if d2.day<d1.day else 0)
    except: pass
    if months==0:
        result="Improcedente"
    elif elapsed is None:
        result="Necessita vistoria"
    elif elapsed<=months:
        result=row["classificacao_padrao"] if row["classificacao_padrao"]!="Necessita vistoria" else "Necessita vistoria"
    else:
        result="Fora do prazo / analisar improcedência"
    return {"resultado":result,"prazo":f"{months} meses" if months else "No ato da entrega",
            "fundamentacao":row["descricao"],"fonte":row["fonte"],"meses_decorridos":elapsed,"regra":dict(row)}

@app.context_processor
def ctx():
    return {"now":datetime.now(),"session":session}

@app.route("/login",methods=["GET","POST"])
def login():
    if request.method=="POST":
        email=request.form.get("email","").strip().lower()
        user=db().execute("SELECT * FROM usuarios WHERE lower(email)=?",(email,)).fetchone()
        if user and check_password_hash(user["senha"],request.form.get("senha","")):
            session.update(user_id=user["id"],nome=user["nome"],perfil=user["perfil"])
            return redirect(url_for("dashboard"))
        flash("E-mail ou senha inválidos.","erro")
    return render_template("login.html")

@app.route("/sair")
def logout():
    session.clear(); return redirect(url_for("login"))

@app.route("/")
@login_required
def dashboard():
    con=db()
    k=con.execute("""SELECT COUNT(*) total,
      SUM(CASE WHEN lower(status) LIKE '%final%' THEN 1 ELSE 0 END) finalizados,
      SUM(CASE WHEN lower(status) NOT LIKE '%final%' THEN 1 ELSE 0 END) abertos,
      SUM(CASE WHEN lower(classificacao) LIKE '%improced%' THEN 1 ELSE 0 END) improcedentes
      FROM chamados""").fetchone()
    por_emp=con.execute("SELECT empreendimento,COUNT(*) qtd FROM chamados GROUP BY empreendimento ORDER BY qtd DESC").fetchall()
    recentes=con.execute("SELECT * FROM chamados ORDER BY id DESC LIMIT 8").fetchall()
    fin=con.execute("""SELECT COALESCE(SUM(CASE WHEN tipo='Entrada' THEN valor ELSE 0 END),0) entradas,
      COALESCE(SUM(CASE WHEN tipo='Saída' THEN valor ELSE 0 END),0) saidas FROM financeiro WHERE status='Pago/Recebido'""").fetchone()
    return render_template("dashboard.html",k=k,por_emp=por_emp,recentes=recentes,fin=fin)

@app.route("/chamados")
@login_required
def chamados():
    q=request.args.get("q","").strip(); emp=request.args.get("empreendimento",""); status=request.args.get("status","")
    sql="SELECT * FROM chamados WHERE 1=1"; args=[]
    if q:
        sql+=" AND (sat LIKE ? OR solicitante LIKE ? OR unidade LIKE ? OR problema LIKE ?)"
        args += [f"%{q}%"]*4
    if emp: sql+=" AND empreendimento=?"; args.append(emp)
    if status: sql+=" AND status=?"; args.append(status)
    sql+=" ORDER BY id DESC"
    con=db()
    rows=con.execute(sql,args).fetchall()
    emps=con.execute("SELECT nome FROM empreendimentos ORDER BY nome").fetchall()
    statuses=con.execute("SELECT DISTINCT status FROM chamados WHERE status<>'' ORDER BY status").fetchall()
    return render_template("chamados.html",rows=rows,emps=emps,statuses=statuses)

@app.route("/chamados/novo",methods=["GET","POST"])
@login_required
def novo_chamado():
    con=db()
    if request.method=="POST":
        categoria=request.form.get("categoria") or guess_category(request.form.get("problema"))
        data_ab=request.form.get("data_recebido") or date.today().isoformat()
        anal=warranty_analysis(request.form.get("empreendimento"),categoria,request.form.get("data_entrega"),data_ab)
        cur=con.execute("""INSERT INTO chamados(sat,empreendimento,data_recebido,solicitante,unidade,problema,
            contato,status,classificacao,categoria,data_entrega,analise_garantia,fundamentacao,responsavel,observacoes)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",(
            request.form.get("sat"),request.form.get("empreendimento"),data_ab,request.form.get("solicitante"),
            request.form.get("unidade"),request.form.get("problema"),request.form.get("contato"),
            request.form.get("status") or "Aberto",request.form.get("classificacao") or anal["resultado"],categoria,
            request.form.get("data_entrega"),anal["resultado"],anal["fundamentacao"],
            request.form.get("responsavel"),request.form.get("observacoes")))
        con.commit()
        flash("Chamado cadastrado com sucesso.","ok")
        return redirect(url_for("ver_chamado",id=cur.lastrowid))
    emps=con.execute("SELECT * FROM empreendimentos ORDER BY nome").fetchall()
    cats=con.execute("SELECT DISTINCT item FROM garantias ORDER BY item").fetchall()
    return render_template("chamado_form.html",c=None,emps=emps,cats=cats)

@app.route("/chamados/<int:id>",methods=["GET","POST"])
@login_required
def ver_chamado(id):
    con=db()
    c=con.execute("SELECT * FROM chamados WHERE id=?",(id,)).fetchone()
    if not c: return "Chamado não encontrado",404
    if request.method=="POST":
        con.execute("""UPDATE chamados SET status=?,classificacao=?,responsavel=?,observacoes=?,categoria=? WHERE id=?""",
                    (request.form.get("status"),request.form.get("classificacao"),request.form.get("responsavel"),
                     request.form.get("observacoes"),request.form.get("categoria"),id))
        con.commit(); flash("Chamado atualizado.","ok"); return redirect(url_for("ver_chamado",id=id))
    fotos=con.execute("SELECT * FROM fotos WHERE chamado_id=? ORDER BY id DESC",(id,)).fetchall()
    mats=con.execute("SELECT * FROM materiais WHERE chamado_id=? ORDER BY id DESC",(id,)).fetchall()
    cats=con.execute("SELECT DISTINCT item FROM garantias ORDER BY item").fetchall()
    return render_template("chamado_detalhe.html",c=c,fotos=fotos,mats=mats,cats=cats)

@app.post("/chamados/<int:id>/analisar")
@login_required
def analisar_chamado(id):
    con=db(); c=con.execute("SELECT * FROM chamados WHERE id=?",(id,)).fetchone()
    categoria=request.form.get("categoria") or guess_category(c["problema"])
    a=warranty_analysis(c["empreendimento"],categoria,c["data_entrega"],c["data_recebido"])
    con.execute("UPDATE chamados SET categoria=?,analise_garantia=?,fundamentacao=? WHERE id=?",
                (categoria,a["resultado"],a["fundamentacao"]+((" Fonte: "+a.get("fonte","")) if a.get("fonte") else ""),id))
    con.commit(); flash("Análise de garantia atualizada.","ok")
    return redirect(url_for("ver_chamado",id=id))

@app.post("/chamados/<int:id>/foto")
@login_required
def foto(id):
    f=request.files.get("foto")
    if not f or not f.filename: flash("Selecione uma foto.","erro"); return redirect(url_for("ver_chamado",id=id))
    name=f"{uuid.uuid4().hex}_{secure_filename(f.filename)}"
    f.save(os.path.join(UPLOAD,name))
    con=db(); con.execute("INSERT INTO fotos(chamado_id,arquivo,descricao) VALUES(?,?,?)",(id,name,request.form.get("descricao")))
    con.commit(); return redirect(url_for("ver_chamado",id=id))

@app.post("/chamados/<int:id>/material")
@login_required
def material(id):
    con=db(); con.execute("""INSERT INTO materiais(chamado_id,item,quantidade,status,observacoes)
                             VALUES(?,?,?,?,?)""",(id,request.form.get("item"),request.form.get("quantidade"),
                             request.form.get("status"),request.form.get("observacoes")))
    con.commit(); return redirect(url_for("ver_chamado",id=id))

@app.route("/garantias")
@login_required
def garantias():
    con=db(); emp=request.args.get("empreendimento",""); q=request.args.get("q","")
    sql="SELECT * FROM garantias WHERE 1=1"; a=[]
    if emp: sql+=" AND empreendimento IN (?, 'Todos')"; a.append(emp)
    if q: sql+=" AND (item LIKE ? OR descricao LIKE ?)"; a += [f"%{q}%",f"%{q}%"]
    sql+=" ORDER BY empreendimento,item,prazo_meses"
    rows=con.execute(sql,a).fetchall()
    return render_template("garantias.html",rows=rows)

@app.route("/financeiro",methods=["GET","POST"])
@login_required
def financeiro():
    con=db()
    if request.method=="POST":
        con.execute("""INSERT INTO financeiro(data,tipo,descricao,empreendimento,categoria,valor,status,forma_pagamento)
                       VALUES(?,?,?,?,?,?,?,?)""",(request.form.get("data"),request.form.get("tipo"),
                       request.form.get("descricao"),request.form.get("empreendimento"),request.form.get("categoria"),
                       float(request.form.get("valor") or 0),request.form.get("status"),request.form.get("forma_pagamento")))
        con.commit(); flash("Lançamento financeiro salvo.","ok"); return redirect(url_for("financeiro"))
    rows=con.execute("SELECT * FROM financeiro ORDER BY data DESC,id DESC").fetchall()
    sums=con.execute("""SELECT COALESCE(SUM(CASE WHEN tipo='Entrada' THEN valor ELSE 0 END),0) entradas,
    COALESCE(SUM(CASE WHEN tipo='Saída' THEN valor ELSE 0 END),0) saidas FROM financeiro WHERE status='Pago/Recebido'""").fetchone()
    emps=con.execute("SELECT nome FROM empreendimentos ORDER BY nome").fetchall()
    return render_template("financeiro.html",rows=rows,sums=sums,emps=emps)

@app.route("/empreendimentos",methods=["GET","POST"])
@login_required
def empreendimentos():
    con=db()
    if request.method=="POST":
        con.execute("INSERT OR IGNORE INTO empreendimentos(nome,data_entrega,observacoes) VALUES(?,?,?)",
                    (request.form.get("nome"),request.form.get("data_entrega"),request.form.get("observacoes")))
        con.commit(); return redirect(url_for("empreendimentos"))
    return render_template("empreendimentos.html",rows=con.execute("SELECT * FROM empreendimentos ORDER BY nome").fetchall())

@app.route("/usuarios",methods=["GET","POST"])
@login_required
@admin_required
def usuarios():
    con=db()
    if request.method=="POST":
        con.execute("INSERT INTO usuarios(nome,email,senha,perfil) VALUES(?,?,?,?)",
                    (request.form.get("nome"),request.form.get("email").lower(),
                     generate_password_hash(request.form.get("senha")),request.form.get("perfil")))
        con.commit(); flash("Usuário cadastrado.","ok"); return redirect(url_for("usuarios"))
    return render_template("usuarios.html",rows=con.execute("SELECT id,nome,email,perfil FROM usuarios ORDER BY nome").fetchall())

@app.route("/exportar")
@login_required
def exportar():
    con=db(); rows=con.execute("SELECT * FROM chamados ORDER BY id").fetchall()
    out=io.StringIO(); w=csv.writer(out,delimiter=";")
    cols=rows[0].keys() if rows else ["id","sat","empreendimento"]
    w.writerow(cols)
    for r in rows: w.writerow([r[c] for c in cols])
    b=io.BytesIO(("\ufeff"+out.getvalue()).encode("utf-8")); b.seek(0)
    return send_file(b,as_attachment=True,download_name="Chamados_BR_Gestao_SAT_8.csv",mimetype="text/csv")

@app.route("/api/analisar",methods=["POST"])
@login_required
def api_analisar():
    d=request.get_json(force=True)
    cat=d.get("categoria") or guess_category(d.get("problema"))
    return jsonify({"categoria":cat,**warranty_analysis(d.get("empreendimento"),cat,d.get("data_entrega"),d.get("data_abertura"))})

if __name__=="__main__":
    app.run(host="0.0.0.0",port=5000,debug=False)
