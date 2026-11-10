# Vendedor 360°

Aplicativo web para apoiar vendedores externos na gestão de clientes, catálogo, preços e pedidos.

## Tecnologias
- React 18 + TypeScript
- Vite
- Lucide React
- GitHub Pages (hospedagem estática)

## Executar localmente
```bash
npm install
npm run dev
```

## Validar e gerar produção
```bash
npm run build
npm run preview
```

## Publicação
O workflow `.github/workflows/deploy.yml` executa o build e publica automaticamente no GitHub Pages após atualizações na branch `main`.

No repositório, acesse **Settings → Pages** e selecione **GitHub Actions** em **Build and deployment → Source**.

## Escopo planejado
1. Login e perfis de acesso.
2. Cadastro e importação de clientes.
3. Cadastro e importação de produtos.
4. Carrinho e cálculo de preço final/descontos.
5. Geração de PDF de orçamento e pedido.
6. Histórico e consulta de pedidos.

## Observação
Esta primeira entrega é a estrutura visual inicial. Os módulos ainda precisam ser implementados e conectados a um banco de dados. Não inserir dados reais de clientes até a implementação de autenticação, banco de dados e controles de segurança.
