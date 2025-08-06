import { useState } from 'react';
import { Button } from './components/ui/button';
import { Input } from './components/ui/input';
import { Label } from './components/ui/label';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from './components/ui/select';
import { Progress } from './components/ui/progress';
import { Upload, Download, ArrowLeft, CheckCircle } from 'lucide-react';
import './styles/globals.css'

import logo from './assets/logo-b.svg';

type FormData = {
  titulo: string;
  professor: string;
  linkedinUrl: string;
  curso: string;
  agradecimento: string;
  arquivo: File | null;
};

type AppState = 'form' | 'processing' | 'completed';

const VITE_BACKEND_URL: string = import.meta.env.VITE_BACKEND_URL as string;

export default function App() {
  const [state, setState] = useState<AppState>('form');
  const [formData, setFormData] = useState<FormData>({
    titulo: '',
    professor: '',
    linkedinUrl: '',
    curso: '',
    agradecimento: '',
    arquivo: null
  });
  const [progress, setProgress] = useState(0);
  const [downloadUrl, setDownloadUrl] = useState<string>('');

  const mbaCourses = [
    { value: "agronegocio", label: "MBA em Agronegócios", name: "Agronegócios", theme: "#5e9440" },
    { value: "controladoria-financas", label: "MBA em Compliance & ESG", name: "Compliance & ESG", theme: "#bb8f12" },
    { value: "economia-mercados", label: "MBA em Data Science e Analytics", name: "Data Science e Analytics", theme: "#8245a4" },
    { value: "finanancas", label: "MBA em Economia, Investimentos e Banking", name: "Economia, Investimentos e Banking", theme: "#bb8f12" },
    { value: "food-business", label: "MBA em Educação Inclusiva e Diversidade", name: "Educação Inclusiva e Diversidade", theme: "#03a6ad" },
    { value: "gestao-comercial", label: "MBA em Engenharia de Software", name: "Engenharia de Software", theme: "#8245a4" },
    { value: "gestao-cooperativas", label: "MBA em ESG e Negócios Sustentáveis", name: "ESG e Negócios Sustentáveis", theme: "#21409a" },
    { value: "gestao-empresarial", label: "MBA em Finanças e Controladoria", name: "Finanças e Controladoria", theme: "#bb8f12" },
    { value: "gestao-inovacao", label: "MBA em Finanças e Valuation", name: "Finanças e Valuation", theme: "#bb8f12" },
    { value: "gestao-pessoas", label: "MBA em Gestão de Negócios", name: "Gestão de Negócios", theme: "#21409a" },
    { value: "gestao-projetos", label: "MBA em Gestão de Negócios Digitais e Inteligência Artificial", name: "Gestão de Negócios Digitais e Inteligência Artificial", theme: "#8245a4" },
    { value: "lideranca-executiva", label: "MBA em Gestão de Pessoas", name: "Gestão de Pessoas", theme: "#21409a" },
    { value: "lideranca-executiva", label: "MBA em Gestão de Projetos", name: "Gestão de Projetos", theme: "#21409a" },
    { value: "lideranca-executiva", label: "MBA em Gestão de Vendas", name: "Gestão de Vendas", theme: "#21409a" },
    { value: "lideranca-executiva", label: "MBA em Gestão Escolar", name: "Gestão Escolar", theme: "#03a6ad" },
    { value: "lideranca-executiva", label: "MBA em Gestão Tributária", name: "Gestão Tributária", theme: "#bb8f12" },
    { value: "lideranca-executiva", label: "MBA em Liderança e Gestão", name: "Liderança e Gestão", theme: "#21409a" },
    { value: "marketing", label: "MBA em Marketing", name: "Marketing", theme: "#21409a" },
    { value: "supply-chain", label: "MBA em Neurociência e Aprendizagem na Educação", name: "Neurociência e Aprendizagem na Educação", theme: "#03a6ad" }
  ];


  const handleInputChange = (field: keyof FormData, value: string) => {
    setFormData(prev => ({ ...prev, [field]: value }));
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0] || null;
    setFormData(prev => ({ ...prev, arquivo: file }));
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!formData.titulo || !formData.professor || !formData.curso || !formData.arquivo) {
      alert('Por favor, preencha todos os campos e selecione um arquivo.');
      return;
    }
    if (!formData.arquivo.name.match(/\.(pptx)$/i)) {
      alert('Por favor, selecione um arquivo válido (.pptx).');
      return;
    }

    setState('processing');
    setProgress(0);
    
    try {
      // Criar FormData para enviar o arquivo
      const formDataToSend = new FormData();
      formDataToSend.append('mba', mbaCourses.find(course => course.value === formData.curso)?.name || formData.curso);
      formDataToSend.append('theme', mbaCourses.find(course => course.value === formData.curso)?.theme || formData.curso);
      formDataToSend.append('tituloAula', formData.titulo);
      formDataToSend.append('nomeProfessor', formData.professor);
      formDataToSend.append('linkedinPerfil', formData.linkedinUrl);
      formDataToSend.append('agradecimento', formData.agradecimento);
      formDataToSend.append('destinationFile', formData.arquivo);

      // Simular progresso enquanto faz a requisição
      const progressInterval = setInterval(() => {
        setProgress(prev => Math.min(prev + 10, 90));
      }, 300);

      // Fazer a requisição para a API
      const response = await fetch(`${VITE_BACKEND_URL}/api/slidemerger/merge`, {
        method: 'POST',
        body: formDataToSend,
      });

      clearInterval(progressInterval);

      if (!response.ok) {
        throw new Error(`Erro na API: ${response.status} - ${response.statusText}`);
      }

      // Verificar se a resposta é um arquivo
      const contentType = response.headers.get('content-type');
      if (contentType && contentType.includes('application/vnd.openxmlformats-officedocument.presentationml.presentation')) {
        // É um arquivo PowerPoint
        const blob = await response.blob();
        const url = URL.createObjectURL(blob);
        setDownloadUrl(url);
      } else {
        // Pode ser uma resposta JSON com URL do arquivo
        const data = await response.json();
        if (data.downloadUrl) {
          setDownloadUrl(data.downloadUrl);
        } else {
          throw new Error('Resposta da API não contém URL de download');
        }
      }

      setProgress(100);
      setState('completed');
    } catch (error) {
      console.error('Erro ao processar arquivo:', error);
      alert(`Erro ao processar arquivo: ${error instanceof Error ? error.message : 'Erro desconhecido'}`);
      setState('form');
      setProgress(0);
    }
  };

  const handleReset = () => {
    setFormData({
      titulo: '',
      professor: '',
      linkedinUrl: '',
      curso: '',
      agradecimento: '',
      arquivo: null
    });
    setProgress(0);
    setDownloadUrl('');
    setState('form');
  };

  const handleDownload = () => {
    if (downloadUrl) {
      const a = document.createElement('a');
      a.href = downloadUrl;
      a.download = formData.arquivo!.name;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
    }
  };

  const isFormValid = formData.titulo && formData.professor && formData.curso && formData.arquivo;

  if (state === 'form') {
    return (
      <>
      <div className="min-h-screen bg-gray-50 flex items-center justify-center px-4 py-8">
        <div className="w-full max-w-lg bg-white rounded-2xl shadow-lg p-8">
          <div className="text-center mb-8">
              <div className="inline-flex items-center justify-center w-16 h-16 rounded-xl mb-4">
                <img src={logo} alt="Logo" className="w-16 h-16" />
              </div>
            <h2 className="text-2xl text-gray-900 mb-2">MBX Standardizer</h2>
            <p className="text-gray-600">Preencha os dados para processar seu arquivo de apresentação</p>
          </div>

          <form onSubmit={handleSubmit} className="space-y-6">
            <div className="space-y-2">
              <Label htmlFor="curso" className="text-gray-700">MBA</Label>
              <Select value={formData.curso} onValueChange={(value) => handleInputChange('curso', value)}>
                <SelectTrigger className="h-12 border-gray-200 focus:border-gray-900 focus:ring-gray-900 rounded-lg">
                  <SelectValue placeholder="Selecione o MBA" />
                </SelectTrigger>
                <SelectContent>
                  {mbaCourses.map((course) => (
                    <SelectItem key={course.value} value={course.value}>
                      {course.label}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>

            <div className="space-y-2">
              <Label htmlFor="titulo" className="text-gray-700">Título da Disciplina</Label>
              <Input
                id="titulo"
                type="text"
                value={formData.titulo}
                onChange={(e) => handleInputChange('titulo', e.target.value)}
                className="h-12 border-gray-200 focus:border-gray-900 focus:ring-gray-900 rounded-lg"
                placeholder="Digite o título da disciplina"
              />
            </div>

            <div className="space-y-2">
              <Label htmlFor="professor" className="text-gray-700">Nome do(a) Professor(a)</Label>
              <Input
                id="professor"
                type="text"
                value={formData.professor}
                onChange={(e) => handleInputChange('professor', e.target.value)}
                className="h-12 border-gray-200 focus:border-gray-900 focus:ring-gray-900 rounded-lg"
                placeholder="Digite o nome do(a) professor(a)"
              />
            </div>

            <div className="space-y-2">
              <Label htmlFor="linkedin" className="text-gray-700">LinkedIn do(a) Professor(a)</Label>
              <Input
                id="linkedin"
                type="url"
                value={formData.linkedinUrl}
                onChange={(e) => handleInputChange('linkedinUrl', e.target.value)}
                className="h-12 border-gray-200 focus:border-gray-900 focus:ring-gray-900 rounded-lg"
                placeholder="https://linkedin.com/in/professor"
              />
            </div>

            <div className="space-y-2">
              <Label htmlFor="agradecimento" className="text-gray-700">Agradecimento</Label>
              <Select value={formData.agradecimento} onValueChange={(value) => handleInputChange('agradecimento', value)}>
                <SelectTrigger className="h-12 border-gray-200 focus:border-gray-900 focus:ring-gray-900 rounded-lg">
                  <SelectValue placeholder="Selecione o Agradecimento" />
                </SelectTrigger>
                <SelectContent>
                    <SelectItem key="agradecimento1" value="1">
                      Obrigado
                    </SelectItem>
                    <SelectItem key="agradecimento2" value="2">
                      Obrigada
                    </SelectItem>
                </SelectContent>
              </Select>
            </div>

            <div className="space-y-2">
              <Label htmlFor="arquivo" className="text-gray-700">Arquivo</Label>
              <div className="border-2 border-dashed border-gray-200 rounded-lg p-6 text-center hover:border-gray-300 transition-colors">
                <Upload className="mx-auto h-8 w-8 text-gray-400 mb-3" />
                <div className="space-y-2">
                  <Input
                    id="arquivo"
                    type="file"
                    onChange={handleFileChange}
                    className="border-0 bg-transparent file:mr-4 file:py-1 file:px-4 file:rounded-lg file:border-0 file:bg-gray-900 file:text-white hover:file:bg-gray-800"
                    accept=".pptx"
                  />
                  {formData.arquivo && (
                    <p className="text-lg text-gray-100 rounded-lg p-2" style={{ backgroundColor: mbaCourses.find(course => course.value === formData.curso)?.theme }}>
                      Arquivo selecionado: {formData.arquivo.name}
                    </p>
                  )}
                </div>
              </div>
            </div>

            <Button
              type="submit"
              className="w-full h-12 bg-gray-900 text-white hover:bg-gray-800 rounded-lg"
              disabled={!isFormValid}
            >
              PROCESSAR ARQUIVO
            </Button>
          </form>

          <div className="mt-6 text-center">
            <p className="text-sm text-gray-500">
              Precisa de ajuda? <a href="#" className="text-gray-900 hover:underline">Entre em contato</a>
            </p>
          </div>
        </div>
      </div>
      </>
    );
  }

  if (state === 'processing') {
    return (
      <div className="min-h-screen flex items-center justify-center bg-gray-50">
        <div className="w-full max-w-md bg-white rounded-2xl shadow-lg p-8">
          <div className="text-center mb-8">
            <div className="inline-flex items-center justify-center w-16 h-16 rounded-xl mb-4">
              <img src={logo} alt="Logo" className="w-16 h-16" />
            </div>
            <h2 className="text-2xl text-gray-900 mb-2">Processando Apresentação</h2>
            <p className="text-gray-600">Aguarde enquanto processamos seu apresentação</p>
          </div>

          <div className="space-y-6">
            <div className="text-center">
              <div className="w-16 h-16 border-4 border-gray-900 border-t-transparent rounded-full animate-spin mx-auto mb-4"></div>
              <Progress value={progress} className="w-full mb-2" />
              <p className="text-sm text-gray-600">{progress}% concluído</p>
            </div>
            
            <div className="bg-gray-50 rounded-lg p-4 space-y-2">
              <div className="flex justify-between text-sm">
                <span className="text-gray-600">Disciplina:</span>
                <span className="text-gray-900">{formData.titulo}</span>
              </div>
              <div className="flex justify-between text-sm">
                <span className="text-gray-600">Professor:</span>
                <span className="text-gray-900">{formData.professor}</span>
              </div>
              <div className="flex justify-between text-sm">
                <span className="text-gray-600">MBA:</span>
                <span className="text-gray-900">
                  {mbaCourses.find(course => course.value === formData.curso)?.label || formData.curso}
                </span>
              </div>
              <div className="flex justify-between text-sm">
                <span className="text-gray-600">Arquivo:</span>
                <span className="text-gray-900">{formData.arquivo?.name}</span>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  if (state === 'completed') {
    return (
      <div className="min-h-screen flex items-center justify-center bg-gray-50">
        <div className="w-full max-w-md bg-white rounded-2xl shadow-lg p-8">
          <div className="text-center mb-8">
            <div className="inline-flex items-center justify-center w-16 h-16 bg-green-100 rounded-xl mb-4">
              <CheckCircle className="h-8 w-8 text-green-600" />
            </div>
            <h2 className="text-2xl text-gray-900 mb-2">Processamento Concluído</h2>
            <p className="text-gray-600">Sua apresentação foi processada com sucesso!</p>
          </div>

          <div className="space-y-6">
            <div className="bg-gray-50 rounded-lg p-4 space-y-2">
              <div className="flex justify-between text-sm">
                <span className="text-gray-600">MBA:</span>
                <span className="text-gray-900">
                  {mbaCourses.find(course => course.value === formData.curso)?.label || formData.curso}
                </span>
              </div>
              <div className="flex justify-between text-sm">
                <span className="text-gray-600">Disciplina:</span>
                <span className="text-gray-900">{formData.titulo}</span>
              </div>
              <div className="flex justify-between text-sm">
                <span className="text-gray-600">Professor:</span>
                <span className="text-gray-900">{formData.professor}</span>
              </div>
              <div className="flex justify-between text-sm">
                <span className="text-gray-600">Arquivo Original:</span>
                <span className="text-gray-900">{formData.arquivo?.name}</span>
              </div>
            </div>

            <div className="space-y-3">
              <Button
                onClick={handleDownload}
                className="w-full h-12 bg-gray-900 text-white hover:bg-gray-800 rounded-lg"
                style={{ backgroundColor: mbaCourses.find(course => course.value === formData.curso)?.theme }}
              >
                <Download className="mr-2 h-4 w-4" />
                BAIXAR ARQUIVO PROCESSADO
              </Button>
              
              <Button
                onClick={handleReset}
                variant="outline"
                className="w-full h-12 border-gray-200 text-gray-900 hover:bg-gray-50 rounded-lg"
              >
                <ArrowLeft className="mr-2 h-4 w-4" />
                PROCESSAR OUTRO ARQUIVO
              </Button>
            </div>
          </div>
        </div>
      </div>
    );
  }

  return null;
}