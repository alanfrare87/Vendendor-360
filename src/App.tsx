import { useEffect, useMemo, useRef, useState } from 'react'
import type { Session } from '@supabase/supabase-js'
import { isSupabaseConfigured, supabase } from './lib/supabase'
import * as XLSX from 'xlsx'
import {
  Activity, ArrowDownRight, ArrowUpRight, BarChart3, Boxes,
  ChevronRight, ClipboardList, FileText, LayoutDashboard, Menu,
  Package, Search, Settings, ShoppingCart, Users, X, Upload,
  FileSpreadsheet, CheckCircle2, AlertCircle
} from 'lucide-react'

type Section = 'Visão geral' | 'Pedidos' | 'Clientes' | 'Produtos' | 'Relatórios' | 'Configurações'
type Client = {
  gv: string; sv: string; rv: string; activity: string; activityName: string;
  cnpj: string; code: string; legalName: string; tradeName: string;
  address: string; cep: string; storeChannel: string; channel: string;
  cluster: string; eligibleEc: string; storePlatform: string;
}
type Product = {
  ean: string
  code: string
  description: string
  price7: number | null
  price28: number | null
  dgManual: number | null
  afterManual: number | null
  ec: number | null
  afterEc: number | null
  pgPlus: number | null
  df: number | null
  dg: number | null
  finalPrice: number | null
}

const STORAGE_KEY = 'vendedor360-products-v1'
const navigation: { label: Section; icon: typeof LayoutDashboard }[] = [
  { label: 'Visão geral', icon: LayoutDashboard },
  { label: 'Pedidos', icon: ClipboardList },
  { label: 'Clientes', icon: Users },
  { label: 'Produtos', icon: Package },
  { label: 'Relatórios', icon: BarChart3 },
  { label: 'Configurações', icon: Settings },
]

const money = (value: number | null) =>
  value === null || !Number.isFinite(value)
    ? '—'
    : value.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })

const percent = (value: number | null) =>
  value === null || !Number.isFinite(value)
    ? '—'
    : `${(value * 100).toLocaleString('pt-BR', { maximumFractionDigits: 2 })}%`

const cellNumber = (value: unknown): number | null => {
  if (value === null || value === undefined || value === '') return null
  const number = typeof value === 'number' ? value : Number(String(value).replace(',', '.'))
  return Number.isFinite(number) ? number : null
}

const cellCode = (value: unknown): string => {
  if (value === null || value === undefined || value === '') return ''
  return String(value).replace(/\.0$/, '').trim()
}

function toClient(row: any): Client {
  return { gv:row.gv_code??'', sv:row.sv_code??'', rv:row.rv_code??'', activity:row.activity??'',
    activityName:row.activity_name??'', cnpj:row.cnpj??'', code:row.client_code??'',
    legalName:row.legal_name??'', tradeName:row.trade_name??'', address:row.address??'',
    cep:row.cep??'', storeChannel:row.store_channel??'', channel:row.channel??'',
    cluster:row.cluster??'', eligibleEc:row.eligible_ec==null?'':String(row.eligible_ec),
    storePlatform:row.store_platform??'' }
}
function toProduct(row: any): Product {
  return { ean:row.ean??'', code:row.product_code??'', description:row.description??'',
    price7:cellNumber(row.price_7_days), price28:cellNumber(row.price_28_days),
    dgManual:cellNumber(row.dg_manual), afterManual:cellNumber(row.price_after_manual),
    ec:cellNumber(row.ec), afterEc:cellNumber(row.price_after_ec), pgPlus:cellNumber(row.pg_plus),
    df:cellNumber(row.df), dg:cellNumber(row.dg), finalPrice:cellNumber(row.final_price) }
}
const chunks = <T,>(items:T[], size:number):T[][] =>
  Array.from({length:Math.ceil(items.length/size)},(_,i)=>items.slice(i*size,(i+1)*size))

function Dashboard({ session, onSignOut }: { session: Session; onSignOut: () => void }) {
  const [section, setSection] = useState<Section>('Visão geral')
  const [menuOpen, setMenuOpen] = useState(false)
  const [products, setProducts] = useState<Product[]>([])
  const [clients, setClients] = useState<Client[]>([])
  const [clientSearch, setClientSearch] = useState('')
  const [clientRole, setClientRole] = useState<'admin' | 'gv' | 'sv' | 'rv'>('rv')
  const [assignedCode, setAssignedCode] = useState('')
  const [isLoadingData, setIsLoadingData] = useState(true)
  const [dataLoadError, setDataLoadError] = useState('')
  const [clientImportMessage, setClientImportMessage] = useState('')
  const [clientImportError, setClientImportError] = useState(false)
  const [isImportingClients, setIsImportingClients] = useState(false)
  const clientFileInput = useRef<HTMLInputElement>(null)
  const [searchTerm, setSearchTerm] = useState('')
  const [importMessage, setImportMessage] = useState('')
  const [importError, setImportError] = useState(false)
  const [isImporting, setIsImporting] = useState(false)
  const fileInput = useRef<HTMLInputElement>(null)

  const filteredClients = useMemo(() => {
    const term = clientSearch.trim().toLocaleLowerCase('pt-BR')
    const searched = term ? clients.filter(client =>
      [client.code, client.cnpj, client.legalName, client.tradeName, client.cep]
        .some(value => value.toLocaleLowerCase('pt-BR').includes(term))
    ) : clients
    return searched.slice(0, 100)
  }, [clients, clientSearch])

  useEffect(() => {
    let cancelled = false
    const loadData = async () => {
      if (!supabase) return
      setIsLoadingData(true)
      setDataLoadError('')
      try {
        const [pr, prod, cli] = await Promise.all([
          supabase.from('profiles').select('role, assigned_code').eq('id', session.user.id).single(),
          supabase.from('products').select('*').eq('active', true).order('description').limit(5000),
          supabase.from('clients').select('*').order('legal_name').limit(5000),
        ])
        if (pr.error) throw pr.error
        if (prod.error) throw prod.error
        if (cli.error) throw cli.error
        if (cancelled) return
        if (['admin','gv','sv','rv'].includes(pr.data.role)) setClientRole(pr.data.role)
        setAssignedCode(pr.data.assigned_code ?? '')
        setProducts((prod.data ?? []).map(toProduct))
        setClients((cli.data ?? []).map(toClient))
      } catch (e) {
        if (!cancelled) setDataLoadError(e instanceof Error ? e.message : 'Falha ao carregar dados do Supabase.')
      } finally {
        if (!cancelled) setIsLoadingData(false)
      }
    }
    void loadData()
    return () => { cancelled = true }
  }, [session.user.id])

  const handleClientImport = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0]
    if (!file) return
    setIsImportingClients(true)
    setClientImportMessage('')
    setClientImportError(false)
    try {
      const buffer = await file.arrayBuffer()
      const workbook = XLSX.read(buffer, { type: 'array', cellText: true, cellDates: false })
      const sheet = workbook.Sheets[workbook.SheetNames[0]]
      if (!sheet) throw new Error('A planilha não possui uma aba válida.')
      const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: '', raw: false })
      const imported: Client[] = rows.slice(1).flatMap(row => {
        const v = row as unknown[]
        const code = cellCode(v[6])
        const cnpj = cellCode(v[5])
        const legalName = cellCode(v[7])
        if (!code && !cnpj && !legalName) return []
        return [{
          gv:cellCode(v[0]), sv:cellCode(v[1]), rv:cellCode(v[2]),
          activity:cellCode(v[3]), activityName:cellCode(v[4]), cnpj, code,
          legalName, tradeName:cellCode(v[8]), address:cellCode(v[9]), cep:cellCode(v[10]),
          storeChannel:cellCode(v[11]), channel:cellCode(v[12]), cluster:cellCode(v[13]),
          eligibleEc:cellCode(v[14]), storePlatform:cellCode(v[15])
        }]
      })
      if (!imported.length) throw new Error('Não encontrei clientes. Confira os cabeçalhos da planilha.')
      if (!supabase) throw new Error('Conexão com o Supabase indisponível.')
      const payload = imported.map(x => ({
        client_code:x.code, gv_code:x.gv||null, sv_code:x.sv||null, rv_code:x.rv||null,
        activity:x.activity||null, activity_name:x.activityName||null, cnpj:x.cnpj||null,
        legal_name:x.legalName||null, trade_name:x.tradeName||null, address:x.address||null,
        cep:x.cep||null, store_channel:x.storeChannel||null, channel:x.channel||null,
        cluster:x.cluster||null,
        eligible_ec:x.eligibleEc===''?null:['true','sim','1','yes'].includes(x.eligibleEc.toLowerCase()),
        store_platform:x.storePlatform||null
      }))
      for (const batch of chunks(payload,500)) {
        const {error}=await supabase.from('clients').upsert(batch,{onConflict:'client_code'})
        if (error) throw error
      }
      const {data,error}=await supabase.from('clients').select('*').order('legal_name').limit(5000)
      if (error) throw error
      setClients((data??[]).map(toClient))
      setClientSearch('')
      setClientImportMessage(imported.length.toLocaleString('pt-BR')+' clientes sincronizados com o Supabase.')
    } catch (error) {
      setClientImportError(true)
      setClientImportMessage(error instanceof Error ? error.message : 'Não foi possível ler a planilha.')
    } finally {
      setIsImportingClients(false)
      event.target.value = ''
    }
  }

  const filteredProducts = useMemo(() => {
    const term = searchTerm.trim().toLocaleLowerCase('pt-BR')
    if (!term) return products
    return products.filter(product =>
      product.description.toLocaleLowerCase('pt-BR').includes(term) ||
      product.code.toLocaleLowerCase('pt-BR').includes(term) ||
      product.ean.includes(term)
    )
  }, [products, searchTerm])

  const handleImport = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0]
    if (!file) return
    setIsImporting(true)
    setImportMessage('')
    setImportError(false)
    try {
      const buffer = await file.arrayBuffer()
      const workbook = XLSX.read(buffer, { type: 'array' })
      const sheet = workbook.Sheets[workbook.SheetNames[0]]
      if (!sheet) throw new Error('A planilha não possui uma aba válida.')
      const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: null, raw: true })
      const imported: Product[] = rows.slice(1).flatMap(row => {
        const values = row as unknown[]
        const code = cellCode(values[1])
        const description = cellCode(values[2])
        if (!code && !description) return []
        return [{
          ean: cellCode(values[0]),
          code,
          description,
          price7: cellNumber(values[3]),
          price28: cellNumber(values[4]),
          dgManual: cellNumber(values[5]),
          afterManual: cellNumber(values[6]),
          ec: cellNumber(values[7]),
          afterEc: cellNumber(values[8]),
          pgPlus: cellNumber(values[9]),
          df: cellNumber(values[10]),
          dg: cellNumber(values[11]),
          finalPrice: cellNumber(values[12]),
        }]
      })
      if (!imported.length) throw new Error('Não encontrei produtos. Confira se a primeira aba contém as colunas EAN, COD e DESCRIÇÃO.')
      if (!supabase) throw new Error('Conexão com o Supabase indisponível.')
      const payload = imported.map(x => ({
        product_code:x.code, ean:x.ean||null, description:x.description,
        price_7_days:x.price7, price_28_days:x.price28, dg_manual:x.dgManual,
        price_after_manual:x.afterManual, ec:x.ec, price_after_ec:x.afterEc,
        pg_plus:x.pgPlus, df:x.df, dg:x.dg, final_price:x.finalPrice, active:true
      }))
      for (const batch of chunks(payload,500)) {
        const {error}=await supabase.from('products').upsert(batch,{onConflict:'product_code'})
        if (error) throw error
      }
      const {data,error}=await supabase.from('products').select('*').eq('active',true).order('description').limit(5000)
      if (error) throw error
      setProducts((data??[]).map(toProduct))
      setSearchTerm('')
      const missingEan = imported.filter(product => !product.ean).length
      setImportMessage(imported.length.toLocaleString('pt-BR')+' produtos sincronizados com o Supabase. '+missingEan+' sem EAN; códigos internos preservados.')
    } catch (error) {
      setImportError(true)
      setImportMessage(error instanceof Error ? error.message : 'Não foi possível ler a planilha.')
    } finally {
      setIsImporting(false)
      event.target.value = ''
    }
  }

  return (
    <div className="app-shell">
      <aside className={`sidebar ${menuOpen ? 'sidebar-open' : ''}`}>
        <div className="brand">
          <div className="brand-mark"><Boxes size={23} /></div>
          <div><strong>VENDEDOR <span>360°</span></strong><small>GESTÃO COMERCIAL</small></div>
          <button className="close-menu" onClick={() => setMenuOpen(false)} aria-label="Fechar menu"><X /></button>
        </div>
        <div className="nav-caption">MENU PRINCIPAL</div>
        <nav>
          {navigation.map(({ label, icon: Icon }) => (
            <button key={label} className={`nav-item ${section === label ? 'active' : ''}`}
              onClick={() => { setSection(label); setMenuOpen(false) }}>
              <Icon size={19} /><span>{label}</span>{section === label && <ChevronRight size={16} />}
            </button>
          ))}
        </nav>
        <div className="sidebar-bottom">
          <div className="profile-avatar">AF</div>
          <div className="profile-info"><strong>Área do vendedor</strong><small>Perfil de demonstração</small></div>
        </div>
      </aside>

      {menuOpen && <button className="overlay" onClick={() => setMenuOpen(false)} aria-label="Fechar menu" />}

      <main className="main-area">
        <header className="topbar">
          <button className="menu-button" onClick={() => setMenuOpen(true)} aria-label="Abrir menu"><Menu /></button>
          <div className="breadcrumb">Vendedor 360° <span>/</span> {section}</div>
          <div className="topbar-right"><span className="status-dot" /> Conectado <button className="signout-button" onClick={onSignOut}>Sair</button></div>
        </header>

        <div className="content">
          <div className="welcome-row">
            <div>
              <p className="eyebrow">PAINEL COMERCIAL</p>
              <h1>{section === 'Visão geral' ? 'Visão geral' : section}</h1>
              <p className="subtitle">Organize sua rotina comercial em um só lugar.</p>
            </div>
            <div className="date-chip"><Activity size={16} /> Vendedor 360°</div>
          </div>

          {section === 'Visão geral' ? (
            <>
              <section className="stats-grid">
                <article className="stat-card"><div className="stat-top"><span>Pedidos no período</span><ClipboardList /></div><strong>0</strong><small><ArrowUpRight size={14} /> Aguardando primeiros pedidos</small></article>
                <article className="stat-card"><div className="stat-top"><span>Clientes cadastrados</span><Users /></div><strong>0</strong><small>Base de clientes</small></article>
                <article className="stat-card"><div className="stat-top"><span>Produtos ativos</span><Package /></div><strong>{products.length.toLocaleString('pt-BR')}</strong><small>Produtos importados no catálogo</small></article>
                <article className="stat-card"><div className="stat-top"><span>Valor em pedidos</span><ShoppingCart /></div><strong>R$ 0,00</strong><small><ArrowDownRight size={14} /> Total registrado</small></article>
              </section>

              <section className="main-grid">
                <article className="panel welcome-panel">
                  <div className="panel-heading"><div><p className="eyebrow">PRIMEIROS PASSOS</p><h2>Bem-vindo ao Vendedor 360°</h2></div><div className="panel-icon"><LayoutDashboard /></div></div>
                  <p className="panel-description">A estrutura inicial do aplicativo está pronta. Vamos evoluir os módulos para transformar visitas em pedidos com agilidade e segurança.</p>
                  <div className="steps">
                    <div className="step"><span className="step-number">01</span><div><strong>Cadastrar clientes</strong><small>Organize sua carteira de atendimento.</small></div><Users /></div>
                    <div className="step"><span className="step-number">02</span><div><strong>Adicionar produtos</strong><small>{products.length ? `${products.length} produtos carregados no catálogo.` : 'Importe a planilha de produtos para começar.'}</small></div><Package /></div>
                    <div className="step"><span className="step-number">03</span><div><strong>Criar primeiro pedido</strong><small>Adicione itens e confira os valores.</small></div><ShoppingCart /></div>
                  </div>
                </article>
                <article className="panel quick-panel">
                  <p className="eyebrow">ACESSO RÁPIDO</p><h2>Comece por aqui</h2>
                  <button className="quick-link" onClick={() => setSection('Pedidos')}><span className="quick-icon"><FileText /></span><span><strong>Novo pedido</strong><small>Montar um pedido de venda</small></span><ChevronRight /></button>
                  <button className="quick-link" onClick={() => setSection('Clientes')}><span className="quick-icon"><Users /></span><span><strong>Clientes</strong><small>Consultar carteira de clientes</small></span><ChevronRight /></button>
                  <button className="quick-link" onClick={() => setSection('Produtos')}><span className="quick-icon"><Package /></span><span><strong>Produtos</strong><small>Consultar catálogo</small></span><ChevronRight /></button>
                  <div className="notice"><Activity size={17} /><span>Clientes e produtos são carregados do Supabase, com acesso controlado pelo perfil.</span></div>
                </article>
              </section>
            </>
          ) : section === 'Clientes' ? (
            <section className="catalog">
              {dataLoadError && <div className="import-message error"><AlertCircle size={17} />{dataLoadError}</div>}
              {isLoadingData && <div className="import-message">Carregando dados do Supabase...</div>}
              <div className="catalog-toolbar">
                <div className="catalog-search"><Search size={18} /><input value={clientSearch} onChange={event => setClientSearch(event.target.value)} placeholder="Pesquisar por cliente, CNPJ, código ou CEP..." /></div>
                <button className="primary-button import-button" onClick={() => clientFileInput.current?.click()} disabled={isImportingClients}><Upload size={16} /> {isImportingClients ? 'Importando...' : 'Importar Excel'}</button>
                <input ref={clientFileInput} className="hidden-file" type="file" accept=".xlsx,.xls,.csv" onChange={handleClientImport} />
              </div>
              <div className="client-access-panel">
                <small>Perfil: {clientRole.toUpperCase()}{assignedCode ? ' • código ' + assignedCode : ''}. O Supabase controla automaticamente sua carteira.</small>
              </div>
              {clientImportMessage && <div className={'import-message ' + (clientImportError ? 'error' : 'success')}>{clientImportError ? <AlertCircle size={17} /> : <CheckCircle2 size={17} />}{clientImportMessage}</div>}
              <div className="catalog-summary"><div><strong>{(clientRole === 'admin' ? clients.length : clients.filter(client => clientRole === 'gv' ? client.gv === assignedCode.trim() : clientRole === 'sv' ? client.sv === assignedCode.trim() : client.rv === assignedCode.trim()).length).toLocaleString('pt-BR')}</strong><span> clientes na carteira</span></div><div><strong>{filteredClients.length.toLocaleString('pt-BR')}</strong><span> exibidos (máx. 100)</span></div></div>
              {isLoadingData ? null : clients.length === 0 ? <div className="panel catalog-empty"><div className="empty-icon"><Users /></div><h2>Importe sua carteira de clientes</h2><p>Carregue a planilha LISTA DE CLIENTES. A busca funciona por código, CNPJ, razão social, nome fantasia e CEP.</p><button className="primary-button" onClick={() => clientFileInput.current?.click()}><Upload size={16} /> Selecionar planilha</button><small>Os registros ficam armazenados no Supabase e respeitam as permissões do seu perfil.</small></div> : <div className="client-list">{filteredClients.length === 0 ? <div className="panel no-results">Nenhum cliente encontrado.</div> : filteredClients.map((client,index) => <article className="client-card" key={client.code + '-' + index}><div className="client-card-head"><strong>{client.tradeName || client.legalName || 'Cliente sem nome'}</strong><span>{client.code}</span></div><p>{client.legalName}</p><div className="client-meta"><span>CNPJ: {client.cnpj}</span><span>CEP: {client.cep || '—'}</span><span>Canal: {client.channel || '—'}</span><span>Cluster: {client.cluster || '—'}</span></div><small>{client.address}</small></article>)}</div>}
            </section>
          ) : section === 'Produtos' ? (
            <section className="catalog">
              <div className="catalog-toolbar">
                <div className="catalog-search"><Search size={18} /><input value={searchTerm} onChange={event => setSearchTerm(event.target.value)} placeholder="Pesquisar por descrição, código ou EAN..." /></div>
                <button className="primary-button import-button" onClick={() => fileInput.current?.click()} disabled={isImporting}>
                  <Upload size={16} /> {isImporting ? 'Importando...' : 'Importar Excel'}
                </button>
                <input ref={fileInput} className="hidden-file" type="file" accept=".xlsx,.xls,.csv" onChange={handleImport} />
              </div>

              {importMessage && <div className={`import-message ${importError ? 'error' : 'success'}`}>{importError ? <AlertCircle size={17} /> : <CheckCircle2 size={17} />}{importMessage}</div>}

              <div className="catalog-summary">
                <div><strong>{products.length.toLocaleString('pt-BR')}</strong><span> produtos no catálogo</span></div>
                <div><strong>{products.filter(product => !product.ean).length.toLocaleString('pt-BR')}</strong><span> sem EAN</span></div>
                <div><strong>{filteredProducts.length.toLocaleString('pt-BR')}</strong><span> exibidos</span></div>
              </div>

              {products.length === 0 ? (
                <div className="panel catalog-empty">
                  <div className="empty-icon"><FileSpreadsheet /></div>
                  <h2>Importe sua tabela de produtos</h2>
                  <p>Selecione a planilha Excel de preços. O sistema preserva preço de 7 e 28 dias, DG manual, EC, P&G+, descontos e preço final calculado.</p>
                  <button className="primary-button" onClick={() => fileInput.current?.click()}><Upload size={16} /> Selecionar planilha</button>
                  <small>Os produtos importados são sincronizados com o catálogo compartilhado no Supabase.</small>
                </div>
              ) : (
                <div className="product-list">
                  {filteredProducts.length === 0 ? <div className="panel no-results">Nenhum produto encontrado para esta pesquisa.</div> :
                    filteredProducts.map((product, index) => (
                      <details className="product-card" key={`${product.code}-${index}`}>
                        <summary>
                          <div className="product-main"><strong>{product.description || 'Produto sem descrição'}</strong><small>Código: {product.code || '—'} <span>•</span> EAN: {product.ean || 'Não informado'}</small></div>
                          <div className="product-price"><small>PREÇO FINAL</small><strong>{money(product.finalPrice)}</strong></div>
                          <ChevronRight className="product-chevron" size={18} />
                        </summary>
                        <div className="price-breakdown">
                          <div><span>Preço 7 dias</span><strong>{money(product.price7)}</strong></div>
                          <div><span>Preço 28 dias</span><strong>{money(product.price28)}</strong></div>
                          <div><span>DG manual</span><strong>{percent(product.dgManual)}</strong></div>
                          <div><span>Após DG manual</span><strong>{money(product.afterManual)}</strong></div>
                          <div><span>Escolha Certa (EC)</span><strong>{percent(product.ec)}</strong></div>
                          <div><span>Após EC</span><strong>{money(product.afterEc)}</strong></div>
                          <div><span>P&G+</span><strong>{percent(product.pgPlus)}</strong></div>
                          <div><span>DF</span><strong>{percent(product.df)}</strong></div>
                          <div><span>DG</span><strong>{percent(product.dg)}</strong></div>
                          <div className="final-breakdown"><span>Preço final calculado</span><strong>{money(product.finalPrice)}</strong></div>
                        </div>
                      </details>
                    ))
                  }
                </div>
              )}
            </section>
          ) : (
            <section className="panel empty-state">
              <div className="empty-icon">{section === 'Pedidos' ? <ShoppingCart /> : section === 'Relatórios' ? <BarChart3 /> : <Settings />}</div>
              <h2>Módulo {section}</h2>
              <p>Esta área já está reservada na estrutura do Vendedor 360°. Vamos implementar suas funções na próxima etapa.</p>
              <button className="primary-button" onClick={() => setSection('Visão geral')}>Voltar ao painel</button>
            </section>
          )}
          <footer>VENDEDOR 360° <span>•</span> Versão inicial 0.4.0</footer>
        </div>
      </main>
    </div>
  )
}

function LoginScreen({ onLogin }: { onLogin: (email: string, password: string) => Promise<void> }) {
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [error, setError] = useState('')
  const [busy, setBusy] = useState(false)

  const submit = async (event: React.FormEvent) => {
    event.preventDefault()
    setBusy(true)
    setError('')
    try {
      await onLogin(email.trim(), password)
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Não foi possível entrar. Confira seus dados.')
    } finally {
      setBusy(false)
    }
  }

  return (
    <main className="auth-screen">
      <form className="auth-card" onSubmit={submit}>
        <div className="auth-brand"><div className="brand-mark"><Boxes size={25} /></div><div><strong>VENDEDOR <span>360°</span></strong><small>GESTÃO COMERCIAL</small></div></div>
        <p className="eyebrow">ACESSO SEGURO</p>
        <h1>Entrar no Vendedor 360°</h1>
        <p className="auth-description">Utilize o e-mail e a senha cadastrados pelo administrador.</p>
        <label htmlFor="login-email">E-mail</label>
        <input id="login-email" type="email" autoComplete="username" value={email} onChange={event => setEmail(event.target.value)} required placeholder="seu@email.com" />
        <label htmlFor="login-password">Senha</label>
        <input id="login-password" type="password" autoComplete="current-password" value={password} onChange={event => setPassword(event.target.value)} required placeholder="Sua senha" />
        {error && <div className="auth-error"><AlertCircle size={17} />{error}</div>}
        <button className="primary-button auth-submit" type="submit" disabled={busy}>{busy ? 'Entrando...' : 'Entrar'}</button>
        <small className="auth-footnote">A criação de usuários é controlada pelo administrador.</small>
      </form>
    </main>
  )
}

function App() {
  const [session, setSession] = useState<Session | null>(null)
  const [loading, setLoading] = useState(true)
  const [configError, setConfigError] = useState('')

  useEffect(() => {
    if (!supabase) {
      setLoading(false)
      return
    }
    supabase.auth.getSession().then(({ data, error }) => {
      if (error) setConfigError(error.message)
      setSession(data.session)
      setLoading(false)
    })
    const { data: { subscription } } = supabase.auth.onAuthStateChange((_event, nextSession) => {
      setSession(nextSession)
    })
    return () => subscription.unsubscribe()
  }, [])

  const login = async (email: string, password: string) => {
    if (!supabase) throw new Error('A conexão com o Supabase não está configurada.')
    const { error } = await supabase.auth.signInWithPassword({ email, password })
    if (error) throw error
  }

  const signOut = async () => {
    if (supabase) await supabase.auth.signOut()
  }

  if (!isSupabaseConfigured || !supabase) {
    return <main className="auth-screen"><section className="auth-card"><h1>Configuração pendente</h1><p>As variáveis do Supabase ainda não foram encontradas nesta publicação. Confira os Secrets do GitHub e execute novamente o workflow.</p></section></main>
  }
  if (loading) return <main className="auth-screen"><section className="auth-card">Conectando ao Vendedor 360°...</section></main>
  if (configError) return <main className="auth-screen"><section className="auth-card"><h1>Falha na conexão</h1><p>{configError}</p><button className="primary-button" onClick={() => window.location.reload()}>Tentar novamente</button></section></main>
  if (!session) return <LoginScreen onLogin={login} />
  return <Dashboard key={session.user.id} session={session} onSignOut={signOut} />
}

export default App