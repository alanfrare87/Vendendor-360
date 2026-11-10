import { useMemo, useRef, useState } from 'react'
import * as XLSX from 'xlsx'
import {
  Activity, ArrowDownRight, ArrowUpRight, BarChart3, Boxes,
  ChevronRight, ClipboardList, FileText, LayoutDashboard, Menu,
  Package, Search, Settings, ShoppingCart, Users, X, Upload,
  FileSpreadsheet, CheckCircle2, AlertCircle
} from 'lucide-react'

type Section = 'Visão geral' | 'Pedidos' | 'Clientes' | 'Produtos' | 'Relatórios' | 'Configurações'
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

function readStoredProducts(): Product[] {
  try {
    const value = localStorage.getItem(STORAGE_KEY)
    return value ? JSON.parse(value) as Product[] : []
  } catch {
    return []
  }
}

function App() {
  const [section, setSection] = useState<Section>('Visão geral')
  const [menuOpen, setMenuOpen] = useState(false)
  const [products, setProducts] = useState<Product[]>(readStoredProducts)
  const [searchTerm, setSearchTerm] = useState('')
  const [importMessage, setImportMessage] = useState('')
  const [importError, setImportError] = useState(false)
  const [isImporting, setIsImporting] = useState(false)
  const fileInput = useRef<HTMLInputElement>(null)

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
      localStorage.setItem(STORAGE_KEY, JSON.stringify(imported))
      setProducts(imported)
      setSearchTerm('')
      const missingEan = imported.filter(product => !product.ean).length
      setImportMessage(`${imported.length} produtos importados. ${missingEan} sem EAN; os códigos internos foram preservados.`)
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
          <div className="topbar-right"><span className="status-dot" /> Ambiente inicial</div>
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
                  <div className="notice"><Activity size={17} /><span>O catálogo fica salvo neste navegador. A sincronização entre vendedores será adicionada com o banco de dados.</span></div>
                </article>
              </section>
            </>
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
                  <small>Os dados serão salvos neste navegador até a integração com o banco de dados.</small>
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
              <div className="empty-icon">{section === 'Pedidos' ? <ShoppingCart /> : section === 'Clientes' ? <Users /> : section === 'Relatórios' ? <BarChart3 /> : <Settings />}</div>
              <h2>Módulo {section}</h2>
              <p>Esta área já está reservada na estrutura do Vendedor 360°. Vamos implementar suas funções na próxima etapa.</p>
              <button className="primary-button" onClick={() => setSection('Visão geral')}>Voltar ao painel</button>
            </section>
          )}
          <footer>VENDEDOR 360° <span>•</span> Versão inicial 0.2.0</footer>
        </div>
      </main>
    </div>
  )
}

export default App