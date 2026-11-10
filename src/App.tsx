import { useState } from 'react'
import {
  Activity, ArrowDownRight, ArrowUpRight, BarChart3, Boxes,
  ChevronRight, ClipboardList, FileText, LayoutDashboard, Menu,
  Package, Search, Settings, ShoppingCart, Users, X
} from 'lucide-react'

type Section = 'Visão geral' | 'Pedidos' | 'Clientes' | 'Produtos' | 'Relatórios' | 'Configurações'

const navigation: { label: Section; icon: typeof LayoutDashboard }[] = [
  { label: 'Visão geral', icon: LayoutDashboard },
  { label: 'Pedidos', icon: ClipboardList },
  { label: 'Clientes', icon: Users },
  { label: 'Produtos', icon: Package },
  { label: 'Relatórios', icon: BarChart3 },
  { label: 'Configurações', icon: Settings },
]

function App() {
  const [section, setSection] = useState<Section>('Visão geral')
  const [menuOpen, setMenuOpen] = useState(false)

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
                <article className="stat-card"><div className="stat-top"><span>Produtos ativos</span><Package /></div><strong>0</strong><small>Catálogo de produtos</small></article>
                <article className="stat-card"><div className="stat-top"><span>Valor em pedidos</span><ShoppingCart /></div><strong>R$ 0,00</strong><small><ArrowDownRight size={14} /> Total registrado</small></article>
              </section>

              <section className="main-grid">
                <article className="panel welcome-panel">
                  <div className="panel-heading"><div><p className="eyebrow">PRIMEIROS PASSOS</p><h2>Bem-vindo ao Vendedor 360°</h2></div><div className="panel-icon"><LayoutDashboard /></div></div>
                  <p className="panel-description">A estrutura inicial do aplicativo está pronta. Vamos evoluir os módulos para transformar visitas em pedidos com agilidade e segurança.</p>
                  <div className="steps">
                    <div className="step"><span className="step-number">01</span><div><strong>Cadastrar clientes</strong><small>Organize sua carteira de atendimento.</small></div><Users /></div>
                    <div className="step"><span className="step-number">02</span><div><strong>Adicionar produtos</strong><small>Monte seu catálogo com preços e códigos.</small></div><Package /></div>
                    <div className="step"><span className="step-number">03</span><div><strong>Criar primeiro pedido</strong><small>Adicione itens e confira os valores.</small></div><ShoppingCart /></div>
                  </div>
                </article>
                <article className="panel quick-panel">
                  <p className="eyebrow">ACESSO RÁPIDO</p><h2>Comece por aqui</h2>
                  <button className="quick-link" onClick={() => setSection('Pedidos')}><span className="quick-icon"><FileText /></span><span><strong>Novo pedido</strong><small>Montar um pedido de venda</small></span><ChevronRight /></button>
                  <button className="quick-link" onClick={() => setSection('Clientes')}><span className="quick-icon"><Users /></span><span><strong>Clientes</strong><small>Consultar carteira de clientes</small></span><ChevronRight /></button>
                  <button className="quick-link" onClick={() => setSection('Produtos')}><span className="quick-icon"><Package /></span><span><strong>Produtos</strong><small>Consultar catálogo</small></span><ChevronRight /></button>
                  <div className="notice"><Activity size={17} /><span>Os dados serão conectados ao banco de dados na próxima etapa.</span></div>
                </article>
              </section>
            </>
          ) : (
            <section className="panel empty-state">
              <div className="empty-icon">{section === 'Pedidos' ? <ShoppingCart /> : section === 'Clientes' ? <Users /> : section === 'Produtos' ? <Package /> : section === 'Relatórios' ? <BarChart3 /> : <Settings />}</div>
              <h2>Módulo {section}</h2>
              <p>Esta área já está reservada na estrutura do Vendedor 360°. Vamos implementar suas funções na próxima etapa.</p>
              <button className="primary-button" onClick={() => setSection('Visão geral')}>Voltar ao painel</button>
            </section>
          )}
          <footer>VENDEDOR 360° <span>•</span> Versão inicial 0.1.0</footer>
        </div>
      </main>
    </div>
  )
}

export default App