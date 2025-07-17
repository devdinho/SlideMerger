# APS - Sistema de Upload Acadêmico

Um sistema simples e elegante para upload e processamento de arquivos acadêmicos, desenvolvido para MBAs da USP Esalq.

## 🚀 Características

- Interface moderna e responsiva
- Upload de arquivos com drag & drop
- Validação de formulários
- Simulação de processamento com barra de progresso
- Download de arquivos processados
- Lista completa de cursos MBA da USP Esalq
- Design clean em preto e branco

## 🛠️ Tecnologias Utilizadas

- **React 18** - Biblioteca para criação de interfaces
- **TypeScript** - Superset JavaScript com tipagem estática
- **Vite** - Build tool e servidor de desenvolvimento
- **Tailwind CSS v4** - Framework CSS utilitário
- **Radix UI** - Componentes primitivos acessíveis
- **Lucide React** - Ícones SVG
- **Shadcn/ui** - Componentes de interface reutilizáveis

## 📦 Instalação

### Pré-requisitos

- Node.js 22+ 
- npm ou yarn

### Passos

1. **Clone ou baixe o projeto**
   ```bash
   # Se usando Git
   git clone <url-do-repositorio>
   cd aps-upload-system
   
   # Ou extraia o arquivo ZIP baixado
   ```

2. **Instale as dependências**
   ```bash
   npm install
   # ou
   yarn install
   ```

3. **Execute o projeto em modo desenvolvimento**
   ```bash
   npm run dev
   # ou
   yarn dev
   ```

4. **Acesse no navegador**
   ```
   http://localhost:5173
   ```

## 🏗️ Scripts Disponíveis

- `npm run dev` - Inicia o servidor de desenvolvimento
- `npm run build` - Cria a build de produção
- `npm run preview` - Visualiza a build de produção
- `npm run lint` - Executa o linter para verificar qualidade do código

## 📁 Estrutura do Projeto

```
src/
├── components/
│   ├── ui/              # Componentes Shadcn/ui
│   └── figma/           # Componentes auxiliares
├── styles/
│   └── globals.css      # Estilos globais e configuração Tailwind
├── App.tsx              # Componente principal
└── main.tsx             # Ponto de entrada da aplicação
```

## 🎯 Funcionalidades

### Formulário de Upload
- Seleção de curso MBA (16 opções disponíveis)
- Campo para título da disciplina
- Campo para nome do professor
- Campo para LinkedIn do professor (opcional)
- Upload de arquivo (aceita .pptx)

### Processamento
- Processamento com barra de progresso
- Exibição dos dados informados
- Feedback visual com spinner animado

### Finalização
- Download do arquivo processado
- Opção para processar novo arquivo
- Feedback de sucesso com ícone

## 🎨 Customização

### Cores e Temas
As cores podem ser personalizadas no arquivo `src/styles/globals.css` através das variáveis CSS customizadas.

### Cursos MBA
Para adicionar ou modificar os cursos MBA, edite o array `mbaCourses` no arquivo `src/App.tsx`.

### Validações
As validações de formulário podem ser encontradas na função `handleSubmit` do componente principal.

## 🚀 Deploy

### Build de Produção
```bash
npm run build
```

Os arquivos de produção serão gerados na pasta `dist/`.

### Opções de Deploy
- **Vercel**: Conecte o repositório e faça deploy automático
- **Netlify**: Arraste a pasta `dist` ou conecte o repositório
- **Servidor próprio**: Sirva os arquivos da pasta `dist`

## 📞 Suporte

Se você encontrar algum problema ou tiver dúvidas:

1. Verifique se todas as dependências foram instaladas corretamente
2. Confirme se está usando Node.js 22+
3. Execute `npm run lint` para verificar problemas no código

## 📄 Licença

Este projeto é privado e destinado ao uso acadêmico na USP Esalq.

---

Desenvolvido com ❤️ para a comunidade acadêmica da USP Esalq