import { Component, signal, computed, OnInit } from '@angular/core';
import * as XLSX from 'xlsx';
import * as FileSaver from 'file-saver';
import jsPDF from 'jspdf';
import autoTable, { Table } from 'jspdf-autotable';
import { ToolbarModule } from 'primeng/toolbar';
import { CommonModule } from '@angular/common';
import { DialogModule } from 'primeng/dialog';
import { TableModule } from 'primeng/table';
import { ButtonModule } from 'primeng/button';
import { InputTextModule } from 'primeng/inputtext';
import { FileUploadModule } from 'primeng/fileupload';
import { TabViewModule } from 'primeng/tabview';
import { DropdownModule } from 'primeng/dropdown';
import { FormsModule } from '@angular/forms';
import { bootstrapApplication, BrowserModule } from '@angular/platform-browser';
import {
  HttpClientModule,
  provideHttpClient,
  withFetch,
} from '@angular/common/http';
import { ToastModule } from 'primeng/toast';
import { MessageService } from 'primeng/api';
import { CurrencyOrDashPipe } from '../../pipes/currency-or-dash.pipe';
import { ValueOrDashPipe } from '../../pipes/value-or-dash.pipe';
import { IndexedDBService } from '../../services/indexeddb.service';

type SistemaImportacao = 'procfit' | 'alpha7';

@Component({
  selector: 'app-home',
  templateUrl: './home.component.html',
  styleUrls: ['./home.component.scss'],
  standalone: true,
  imports: [
    HttpClientModule,
    CommonModule,
    FormsModule,
    DialogModule,
    FileUploadModule,
    ButtonModule,
    ToolbarModule,
    TableModule,
    InputTextModule,
    TabViewModule,
    DropdownModule,
    ToastModule,
    CurrencyOrDashPipe,
    ValueOrDashPipe,
  ],
  providers: [MessageService],
})
export class HomeComponent implements OnInit {
  // Permite editar a quantidade escaneada direto na tabela
  onQtdeEscaneadaChange(item: any, value: string) {
    const novaQtde = Number(value);
    if (!isNaN(novaQtde) && novaQtde >= 0) {
      const lista = this.contagemDetalhada();
      const idx = lista.findIndex((i) => i.codigo === item.codigo);
      if (idx !== -1) {
        lista[idx].qtdeEscaneada = novaQtde;
        this.contagemDetalhada.set([...lista]);
        this.gerarDivergencias(); // ADICIONADO
        // Persistir contagem ao editar quantidade
        this.indexedDBService
          .setItem('contagemDetalhada', this.contagemDetalhada())
          .catch(() => {});
      }
    }
  }

  showCadastroDialog = signal(false);
  showContagemDialog = signal(false);
  cadastro = signal<any[]>([]);
  sistemasImportacao: { label: string; value: SistemaImportacao }[] = [
    { label: 'PROCFIT', value: 'procfit' },
    { label: 'ALPHA 7 - ', value: 'alpha7' },
  ];
  sistemaSelecionado = signal<SistemaImportacao>('procfit');
  isAlpha7 = computed(() => this.sistemaSelecionado() === 'alpha7');
  acceptCadastroExtensions = computed(() =>
    this.isAlpha7() ? '.txt' : '.xls,.xlsx'
  );
  divergencias = signal<any[]>([]);
  sistemaSelecionadoModel: SistemaImportacao = 'procfit';

  contagemDetalhada = signal<any[]>([]);
  naoInventariados = signal<any[]>([]);
  eanColumns = computed(() => {
    const produtos = this.cadastro();
    const max = produtos.reduce(
      (acc, p) => Math.max(acc, this.splitEans(p.ean).length),
      0
    );
    // Se não houver nenhum EAN em nenhum produto, retorna array vazia
    return Array.from({ length: Math.min(12, max) }, (_, i) => i);
  });
  // Colunas EAN por aba/dataset
  eanColumnsCadastro = computed(() => {
    const produtos = this.cadastro();
    const max = produtos.reduce((acc, p) => Math.max(acc, this.splitEans(p.ean).length), 0);
    return Array.from({ length: Math.min(12, max) }, (_, i) => i);
  });
  eanColumnsContagem = computed(() => {
    const itens = this.contagemDetalhada();
    const max = itens.reduce((acc, p) => Math.max(acc, this.splitEans(p.ean).length), 0);
    return Array.from({ length: Math.min(12, max) }, (_, i) => i);
  });
  eanColumnsDivergencias = computed(() => {
    const itens = this.divergencias();
    const max = itens.reduce((acc, p) => Math.max(acc, this.splitEans(p.ean).length), 0);
    return Array.from({ length: Math.min(12, max) }, (_, i) => i);
  });
  eanColumnsNaoContados = computed(() => {
    const itens = this.naoInventariados();
    const max = itens.reduce((acc, p) => Math.max(acc, this.splitEans(p.ean).length), 0);
    return Array.from({ length: Math.min(12, max) }, (_, i) => i);
  });
  footerText =
    'Razão Social:EB  Inventário CNPJ: 51.390.090/0001-05 Telefone: (11)99958-5344 Endereço: Rua América Central, 285, Parque das Américas, Mauá - SP, 09351-190';

  private indexedDBService = new IndexedDBService('GestaoBalancoDB', 'CadastroProdutos');

  constructor(private messageService: MessageService) { }
  async ngOnInit() {
    this.sistemaSelecionadoModel = this.sistemaSelecionado();
    try {
      const cadastroProdutos = await this.indexedDBService.getItem('cadastroProdutos');
      if (Array.isArray(cadastroProdutos) && cadastroProdutos.length) {
        const normalizados = cadastroProdutos.map((p: any) => ({
          ...p,
          ean: Array.isArray(p?.ean)
            ? p.ean.filter(Boolean).join(';')
            : (p?.ean ?? '').toString().trim(),
        }));
        this.cadastro.set(normalizados);
      }

      const contagemSalva = await this.indexedDBService.getItem('contagemDetalhada');
      if (Array.isArray(contagemSalva) && contagemSalva.length) {
        this.contagemDetalhada.set(contagemSalva);
      }

      if (this.cadastro().length || this.contagemDetalhada().length) {
        this.gerarDivergencias();
      }
    } catch (error) {
      console.error('Falha ao carregar dados do IndexedDB no boot:', error);
    }
  }

  async limparDados() {
    try {
      await this.indexedDBService.clear();
      this.cadastro.set([]);
      this.contagemDetalhada.set([]);
      this.divergencias.set([]);
      this.naoInventariados.set([]);
      this.messageService.add({
        severity: 'success',
        summary: 'Dados limpos',
        detail: 'Todos os dados locais foram removidos.'
      });
    } catch (error) {
      console.error('Erro ao limpar IndexedDB:', error);
      this.messageService.add({
        severity: 'error',
        summary: 'Erro ao limpar dados',
        detail: 'Não foi possível limpar os dados locais.'
      });
    }
  }

  onSistemaSelecionado(valor: SistemaImportacao | null | undefined) {
    if (!valor) {
      return;
    }
    this.sistemaSelecionadoModel = valor;
    this.sistemaSelecionado.set(valor);
  }

  gerarDivergencias() {
    const contagem = this.contagemDetalhada();
    const cadastro = this.cadastro();


    const cadastroMap = new Map(cadastro.map(p => [String(p.codigo).replace(/^0+/, ''), p]));
    const contagemMap = new Map(contagem.map(c => [String(c.codigo).replace(/^0+/, ''), c]));


    const divergencias: any[] = [];
    const naoContados: any[] = []; // NOVO


    contagem.forEach(c => {
      const cod = String(c.codigo).replace(/^0+/, '');
      const cad = cadastroMap.get(cod);


      if (!cad) {
        divergencias.push({ tipo: 'Produto não cadastrado', ...c });
      } else {
        const qtdeLoja = Number(cad.qtde ?? 0);
        const qtdeEscaneada = Number(c.qtdeEscaneada ?? 0);
        const qtdeDivergente = qtdeEscaneada - qtdeLoja;
        const custo = Number(cad.custo ?? 0);
        const valorDiferenca = qtdeDivergente * custo;


        if (qtdeDivergente !== 0) {
          divergencias.push({
            ...cad,
            qtdeEscaneada,
            qtdeLoja,
            qtdeDivergente,
            valorDiferenca,
          });
        }
      }
    });


    // Verifica quais produtos do cadastro não estão na contagem
    cadastro.forEach(cad => {
      const cod = String(cad.codigo).replace(/^0+/, '');
      if (!contagemMap.has(cod) && Number(cad.qtde ?? 0) > 0) {
        naoContados.push(cad);
      }
    });

    // Também incluir na divergência os produtos não contados com saldo (>0)
    naoContados.forEach((cad) => {
      const qtdeLoja = Number(cad.qtde ?? 0);
      const qtdeEscaneada = 0;
      const qtdeDivergente = qtdeEscaneada - qtdeLoja; // negativo
      const custo = Number(cad.custo ?? 0);
      const valorDiferenca = qtdeDivergente * custo;
      divergencias.push({
        ...cad,
        tipo: 'Não contado com estoque',
        qtdeEscaneada,
        qtdeLoja,
        qtdeDivergente,
        valorDiferenca,
      });
    });

    this.divergencias.set(divergencias);
    this.naoInventariados.set(naoContados);
  }

  // Importa cadastro conforme o sistema selecionado
  handleCadastroUpload(event: any) {
    try {
      const file: File | undefined = event?.files?.[0];
      if (!file) {
        this.messageService.add({
          severity: 'warn',
          summary: 'Nenhum arquivo selecionado',
          detail: 'Selecione um arquivo para importar.',
        });
        return;
      }

      const sistema = this.sistemaSelecionado();
      const leitor =
        sistema === 'alpha7'
          ? this.lerCadastroAlpha7(file)
          : this.lerCadastroProcfit(file);

      leitor
        .then((cadastroProdutos) =>
          this.persistirCadastroProdutos(cadastroProdutos, sistema)
        )
        .catch((error) => {
          const alreadyHandled = !!(
            error && typeof error === 'object' && (error as any).__handled
          );
          if (alreadyHandled) {
            return;
          }
          console.error('Importacao de cadastro:', error);
          this.messageService.add({
            severity: 'error',
            summary: 'Erro ao importar cadastro',
            detail:
              typeof error === 'string'
                ? error
                : 'Verifique o formato do arquivo selecionado.',
          });
        });
    } catch (error) {
      console.error('Importacao de cadastro:', error);
      this.messageService.add({
        severity: 'error',
        summary: 'Erro ao importar cadastro',
        detail: 'Nao foi possivel ler o arquivo selecionado.',
      });
    }
  }

  private lerCadastroProcfit(file: File): Promise<any[]> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e: any) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          const worksheet = workbook.Sheets[workbook.SheetNames[0]];
          const rawData = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

          const normalizeKey = (key: string) =>
            key.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase();

          const toStringTrim = (value: any) =>
            value === undefined || value === null ? '' : String(value).trim();

          const cadastroProdutos = (rawData || []).map((row: any) => {
            const normalized: any = {};
            for (const k of Object.keys(row)) {
              normalized[normalizeKey(k)] = row[k];
            }

            const eanValues: string[] = [];
            const toPush = (val: any) => {
              const parts = this.splitEans(toStringTrim(val));
              if (parts && parts.length) eanValues.push(...parts);
            };
            const explicitKeys = ['EAN', 'CODIGO_EAN', 'CODIGO DE BARRAS', 'COD_BARRAS'];
            for (const key of explicitKeys) {
              if (normalized[key] !== undefined && normalized[key] !== '') {
                toPush(normalized[key]);
              }
            }
            Object.keys(normalized).forEach((key) => {
              const m = key.match(/^EAN\s*[_-]?\s*(\d+)?$/);
              if (m) {
                const idx = m[1] ? parseInt(m[1], 10) : 1;
                if (!Number.isNaN(idx) && idx >= 1 && idx <= 12) {
                  toPush(normalized[key]);
                }
              }
            });
            const eansDedup = Array.from(new Set(eanValues.filter(Boolean)));

            return {
              codigo: toStringTrim(
                normalized['PRODUTO'] ?? normalized['CODIGO'] ?? normalized['ID']
              ),
              ean: eansDedup.join(';'),
              descricao: toStringTrim(normalized['DESCRICAO']),
              fabricante: toStringTrim(normalized['EMPRESA']),
              qtde: Number(normalized['SALDO'] ?? normalized['QTDE'] ?? 0),
              controlado: toStringTrim(normalized['CONTROLADO']),
              custo: this.converterParaFloat(
                normalized['CUSTO_GERENCIAL'] ?? normalized['CUSTO']
              ),
              secao: toStringTrim(normalized['SECAO']),
            };
          });

          resolve(cadastroProdutos);
        } catch (err) {
          reject(err);
        }
      };
      reader.onerror = () => reject('Falha ao ler o arquivo XLS/XLSX.');
      reader.readAsArrayBuffer(file);
    });
  }

  private lerCadastroAlpha7(file: File): Promise<any[]> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e: any) => {
        try {
          const conteudo = (e.target.result as string) ?? '';
          const linhas = conteudo
            .split(/\r?\n/)
            .map((l) => l.trim())
            .filter((l) => l.length);

          const cadastroProdutos = linhas
            .map((linha, index) => {
              if (index === 0) {
                return null;
              }
              if (linha.toLowerCase().startsWith('etiqueta;')) {
                return null;
              }
              const partes = linha.split(';');
              if (partes.length < 7) {
                console.warn(`Linha ${index + 1} ignorada: formato invalido.`, linha);
                return null;
              }
              const [
                etiqueta,
                produto,
                codigoBarras,
                fabricante,
                custo,
                estoque,
                controlado,
              ] = partes.map((p) => p.trim());

              const eans = Array.from(
                new Set(this.splitEans(codigoBarras).filter(Boolean))
              );

              return {
                codigo: etiqueta,
                ean: eans.join(';'),
                descricao: produto,
                fabricante,
                qtde: this.converterParaFloat(estoque),
                controlado: (controlado || '').toUpperCase(),
                custo: this.converterParaFloat(custo),
                secao: '',
              };
            })
            .filter((item): item is any => !!item);

          resolve(cadastroProdutos);
        } catch (err) {
          reject(err);
        }
      };
      reader.onerror = () => reject('Falha ao ler o arquivo TXT.');
      reader.readAsText(file, 'ISO-8859-1');
    });
  }

  private persistirCadastroProdutos(cadastroProdutos: any[], sistema: SistemaImportacao) {
    const total = Array.isArray(cadastroProdutos) ? cadastroProdutos.length : 0;

    return this.indexedDBService
      .setItem('cadastroProdutos', cadastroProdutos)
      .then(() => this.indexedDBService.getItem('cadastroProdutos'))
      .then((produtos) => {
        const normalizados = (produtos || []).map((p: any) => ({
          ...p,
          ean: Array.isArray(p?.ean)
            ? p.ean.filter(Boolean).join(';')
            : (p?.ean ?? '').toString().trim(),
        }));
        this.cadastro.set(normalizados);

        this.messageService.add({
          severity: 'success',
          summary: 'Produtos carregados',
          detail: `${normalizados.length} produtos recuperados do IndexedDB (${this.descricaoSistemaImportacao(sistema)}).`,
        });

        this.contagemDetalhada.set([]);
        this.divergencias.set([]);
        this.naoInventariados.set([]);
        this.indexedDBService.setItem('contagemDetalhada', []).catch(() => {});

        this.messageService.add({
          severity: total > 0 ? 'success' : 'warn',
          summary: total > 0 ? 'Cadastro importado com sucesso' : 'Cadastro vazio',
          detail:
            total > 0
              ? `${total} produtos carregados (${this.descricaoSistemaImportacao(sistema)}).`
              : 'Nenhum produto foi encontrado no arquivo importado.',
        });

        this.showCadastroDialog.set(false);
      })
      .catch((error) => {
        this.messageService.add({
          severity: 'error',
          summary: 'Erro ao salvar/carregar cadastro',
          detail: 'Nao foi possivel salvar ou recuperar os produtos do IndexedDB.',
        });
        console.error('IndexedDB:', error);
        const handledError: any =
          error && typeof error === 'object' ? error : new Error(String(error));
        handledError.__handled = true;
        throw handledError;
      });
  }

  private descricaoSistemaImportacao(sistema: SistemaImportacao): string {
    return sistema === 'alpha7' ? 'Alpha 7' : 'Procfit';
  }

  converterParaFloat(valor: any): number {
    if (!valor) return 0;
    return parseFloat(
      String(valor)
        .replace(/[R$\s]/g, '')
        .replace(',', '.')
    ) || 0;
  }

  // Quebra a célula em array de EANs, aceitando ; , quebra de linha e removendo vazios
  public splitEans(value: any): string[] {
    if (Array.isArray(value)) return value.map(v => String(v).trim()).filter(Boolean);
    if (value === undefined || value === null) return [];
    return String(value)
      .split(/[;,\n\r]+/)
      .map(v => v.trim())
      .filter(Boolean);
  }

  // Retorna os índices de colunas EAN que possuem ao menos um valor entre os itens fornecidos
  public getEanIndexesWithData(items: any[]): number[] {
    const arr = Array.isArray(items) ? items : [];
    const maxLen = Math.min(
      12,
      arr.reduce((acc, p) => Math.max(acc, this.splitEans(p?.ean).length), 0)
    );
    if (maxLen === 0) return [];
    const present: boolean[] = Array(maxLen).fill(false);
    for (const p of arr) {
      const eans = this.splitEans(p?.ean);
      const limit = Math.min(maxLen, eans.length);
      for (let i = 0; i < limit; i++) {
        if (eans[i] && String(eans[i]).trim()) present[i] = true;
      }
    }
    return present.map((ok, i) => (ok ? i : -1)).filter((i) => i >= 0);
  }

  // Calcula quantas colunas de EAN devem existir (máximo entre todas as listas)
  // Removido método duplicado eanColumns() para evitar erro de identificador duplicado.
  handleContagemUpload(event: any) {
    try {
      const files = event.files as File[];

      const normalizarCodigo = (v: any) =>
        String(v ?? '').trim().replace(/^0+/, '');

      // Acumulador por código
      const contagemMap = new Map<string, any>();

      this.indexedDBService.getItem('cadastroProdutos').then((cadastroProdutos) => {
        // Indexar por código normalizado e por cada EAN individual
        const codigoMap = new Map(
          (cadastroProdutos || []).map((p: any) => [normalizarCodigo(p.codigo), p])
        );
        const normalizarEan = (v: any) => String(v ?? '').trim().replace(/\D/g, '');
        const eanMap = new Map<string, any>();
        (cadastroProdutos || []).forEach((p: any) => {
          this.splitEans(p.ean).forEach((e) => {
            const key = normalizarEan(e);
            if (key) eanMap.set(key, p);
          });
        });

        const processarArquivo = (file: File) =>
          new Promise<void>((resolve) => {
            const reader = new FileReader();
            reader.onload = (e: any) => {
              (e.target.result as string)
                .split(/\r?\n/)
                .filter((l) => l.trim().length)
                .forEach((l) => {
                  const [cod, qtde, secao] = l.split(/[|,;]/);
                  const raw = (cod ?? '').trim();
                  const keyCodigo = normalizarCodigo(raw);
                  const keyEan = normalizarEan(raw);

                  const prod = codigoMap.get(keyCodigo) ?? eanMap.get(keyEan);
                  const semCadastro = !prod;
                  const contKey = prod ? normalizarCodigo(prod.codigo) : (keyCodigo || keyEan);

                  let atual = contagemMap.get(contKey);
                  if (!atual) {
                    atual = {
                      codigo: prod?.codigo ?? raw,
                      ean: prod?.ean ?? raw,
                      descricao: semCadastro
                        ? 'PRODUTO NÃO ENCONTRADO NO CADASTRO'
                        : (prod?.descricao ?? ''),
                      fabricante: semCadastro ? '---' : (prod?.fabricante ?? '---'),
                      secao: semCadastro ? '---' : '',
                      qtdeEscaneada: 0,
                      semCadastro,
                    };
                  } else {
                    atual.semCadastro = atual.semCadastro || semCadastro;
                  }

                  if (semCadastro || atual.semCadastro) {
                    atual.semCadastro = true;
                    atual.descricao = 'PRODUTO NÃO ENCONTRADO NO CADASTRO';
                    atual.fabricante = '---';
                    atual.secao = '---';
                  } else if (prod) {
                    atual.codigo = prod.codigo ?? atual.codigo;
                    atual.ean = prod.ean ?? atual.ean;
                    atual.descricao = prod.descricao ?? atual.descricao ?? '';
                    atual.fabricante = prod.fabricante ?? atual.fabricante ?? '---';
                    const novaSecao = (secao ?? '').trim() || prod.secao || '';
                    if (novaSecao) atual.secao = novaSecao;
                    atual.semCadastro = false;
                  }

                  atual.qtdeEscaneada += Number(qtde) || 0;
                  contagemMap.set(contKey, atual);
                });

              resolve();
            };
            reader.readAsText(file);
          });

        Promise.all(files.map(processarArquivo)).then(() => {
          const lista = Array.from(contagemMap.values());
          this.contagemDetalhada.set(lista);
          this.gerarDivergencias();
          // Persistir contagem importada e fechar o modal
          this.indexedDBService.setItem('contagemDetalhada', lista).catch(() => {});
          this.messageService.add({
            severity: 'success',
            summary: 'Contagem importada com sucesso',
            detail: `${lista.length} produtos processados e vinculados ao cadastro.`,
          });
          this.showContagemDialog.set(false);
        });
      }).catch((error) => {
        this.messageService.add({
          severity: 'error',
          summary: 'Erro ao carregar cadastro',
          detail: 'Não foi possível recuperar os produtos do IndexedDB.',
        });
        console.error('Erro ao acessar IndexedDB:', error);
      });
    } catch (error) {
      this.messageService.add({
        severity: 'error',
        summary: 'Erro ao importar contagem',
        detail: 'Verifique o conteúdo dos arquivos TXT.',
      });
      console.error('Erro ao processar contagem:', error);
    }
  }


  exportarPdf() {
    const doc = new jsPDF({ orientation: 'landscape' });

    // Título
    doc.setFontSize(16);
    doc.setFont('helvetica', 'bold');
    doc.text('RELATÓRIO DE DIVERGÊNCIA', doc.internal.pageSize.getWidth() / 2, 15, { align: 'center' });

    // Tabela
    autoTable(doc, {
      startY: 25,
      head: [
        [
          'Empresa',
          'Código',
          'EAN',
          'Descrição',
          'Seção',
          'Custo',
          'Qtde Loja',
          'Qtde Escaneada',
          'Qtde Divergente',
          'Valor Diferença',
        ],
      ],
      body: this.divergencias()
        .slice()
        .sort((a, b) =>
          (a.secao ?? '').localeCompare(b.secao ?? '') ||
          a.descricao.localeCompare(b.descricao)
        )
        .map((d) => [

          d.fabricante ?? '---',
          d.codigo ?? '---',
          (this.splitEans(d.ean)[0] ?? '---'),
          d.descricao ?? '---',
          d.secao ?? '---',
          d.custo == null || isNaN(d.custo)
            ? '---'
            : d.custo.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }),
          d.qtdeLoja ?? '---',
          d.qtdeEscaneada ?? '---',
          d.qtdeDivergente ?? '---',
          d.valorDiferenca == null || isNaN(d.valorDiferenca)
            ? '---'
            : d.valorDiferenca.toLocaleString('pt-BR', {
              style: 'currency',
              currency: 'BRL',
            }),
        ]),
    });

    const finalY = (doc as any).lastAutoTable.finalY || 25;

    doc.setFontSize(12);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(0, 0, 255);
    doc.text(
      `Divergências Positivas: ${this.totalPositivo.toLocaleString('pt-BR', {
        style: 'currency',
        currency: 'BRL',
      })}`,
      14,
      finalY + 10
    );

    doc.setTextColor(255, 0, 0);
    doc.text(
      `Divergências Negativas: ${this.totalNegativo.toLocaleString('pt-BR', {
        style: 'currency',
        currency: 'BRL',
      })}`,
      14,
      finalY + 15
    );

    const totalColor = this.diferencaBalanco >= 0 ? [0, 0, 255] : [255, 0, 0];
    doc.setTextColor(totalColor[0], totalColor[1], totalColor[2]);
    doc.text(
      `Total Divergências: ${this.diferencaBalanco.toLocaleString('pt-BR', {
        style: 'currency',
        currency: 'BRL',
      })}`,
      14,
      finalY + 20
    );

    const pageHeight = doc.internal.pageSize.getHeight();
    const pageWidth = doc.internal.pageSize.getWidth();
    doc.setFontSize(10);
    doc.setTextColor(0, 0, 0);
    doc.setFont('helvetica', 'normal');
    doc.text(this.footerText, pageWidth / 2, pageHeight - 10, { align: 'center' });

    doc.save('divergencias.pdf');
  }


  exportarExcel() {
    const dados = this.divergencias()
      .slice()
      .sort((a, b) =>
        (a.secao ?? '').localeCompare(b.secao ?? '') ||
        a.descricao.localeCompare(b.descricao)
      )
      .map((d) => ({
        codigo: d.codigo ?? '---',
        ean: d.ean ?? '---',
        descricao: d.descricao ?? '---',
        fabricante: d.fabricante ?? '---',
        secao: d.secao ?? '---',
        custo: d.custo == null || isNaN(d.custo) ? '---' : d.custo,
        qtdeLoja: d.qtdeLoja ?? '---',
        qtdeEscaneada: d.qtdeEscaneada ?? '---',
        qtdeDivergente: d.qtdeDivergente ?? '---',
        valorDiferenca:
          d.valorDiferenca == null || isNaN(d.valorDiferenca)
            ? '---'
            : d.valorDiferenca,
      }));
    const ws = XLSX.utils.json_to_sheet(dados, {
      header: [
        'codigo',
        'ean',
        'descricao',
        'fabricante',
        'secao',
        'custo',
        'qtdeLoja',
        'qtdeEscaneada',
        'qtdeDivergente',
        'valorDiferenca',
      ],
    });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Divergencias');
    const excelBuffer: any = XLSX.write(wb, {
      bookType: 'xlsx',
      type: 'array',
    });
    const data: Blob = new Blob([excelBuffer], {
      type: 'application/octet-stream',
    });
    FileSaver.saveAs(data, 'Relatório de divergências.xlsx');
  }

  get totalPositivo(): number {
    return this.divergencias()
      .filter((d) => d.valorDiferenca != null && d.valorDiferenca > 0)
      .reduce((sum, d) => sum + d.valorDiferenca, 0);
  }

  get totalNegativo(): number {
    return this.divergencias()
      .filter((d) => d.valorDiferenca != null && d.valorDiferenca < 0)
      .reduce((sum, d) => sum + d.valorDiferenca, 0);
  }

  get diferencaBalanco(): number {
    return this.totalPositivo + this.totalNegativo;
  }

  onArquivoSelecionado(event: any): void {
    console.log('Arquivo selecionado:', event);
  }

  uploadEvent(callback: Function) {
    callback();
  }

  choose(event: Event, callback: Function) {
    callback();
  }

  onSelectedFiles(event: any) {
    console.log('Arquivos selecionados:', event.files);
  }

  formatSize(bytes: number): string {
    if (bytes === 0) return '0 Bytes';
    const k = 1024,
      sizes = ['Bytes', 'KB', 'MB', 'GB'],
      i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  }

  exportarCadastroPdf() {
    const doc = new jsPDF({ orientation: 'landscape' });

    doc.setFontSize(16);
    doc.setFont('helvetica', 'bold');
    doc.text('CADASTRO DE PRODUTOS', doc.internal.pageSize.getWidth() / 2, 15, { align: 'center' });

    const isAlpha = this.isAlpha7();
    const head = isAlpha
      ? [['ETIQUETA', 'DESCRICAO', 'EAN', 'SALDO', 'CONTROLADO', 'CUSTO_GERENCIAL']]
      : [['EMPRESA', 'PRODUTO', 'DESCRICAO', 'EAN', 'SALDO', 'CONTROLADO', 'CUSTO_GERENCIAL']];
    const formatarCusto = (valor: any) =>
      valor == null || isNaN(valor)
        ? '---'
        : Number(valor).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });

    autoTable(doc, {
      startY: 25,
      head,
      body: this.cadastro()
        .slice()
        .sort((a, b) => a.descricao.localeCompare(b.descricao))
        .map((p) =>
          isAlpha
            ? [
                p.codigo,
                p.descricao,
                this.splitEans(p.ean)[0] ?? '',
                p.qtde,
                p.controlado,
                formatarCusto(p.custo),
              ]
            : [
                p.fabricante,
                p.codigo,
                p.descricao,
                p.ean,
                p.qtde,
                p.controlado,
                formatarCusto(p.custo),
              ]
        ),
    });

    const pageHeight = doc.internal.pageSize.getHeight();
    const pageWidth = doc.internal.pageSize.getWidth();
    doc.setFontSize(10);
    doc.setTextColor(0, 0, 0);
    doc.setFont('helvetica', 'normal');
    doc.text(this.footerText, pageWidth / 2, pageHeight - 10, { align: 'center' });

    doc.save('cadastro.pdf');
  }

  exportarCadastroExcel() {
    const isAlpha = this.isAlpha7();
    const dados = this.cadastro()
      .slice()
      .sort((a, b) => a.descricao.localeCompare(b.descricao))
      .map((p) =>
        isAlpha
          ? {
              ETIQUETA: p.codigo,
              DESCRICAO: p.descricao,
              EAN: this.splitEans(p.ean)[0] ?? '',
              SALDO: p.qtde,
              CONTROLADO: p.controlado,
              CUSTO_GERENCIAL: p.custo == null || isNaN(p.custo) ? '---' : p.custo,
            }
          : {
              EMPRESA: p.fabricante,
              PRODUTO: p.codigo,
              DESCRICAO: p.descricao,
              EAN: p.ean,
              SALDO: p.qtde,
              CONTROLADO: p.controlado,
              CUSTO_GERENCIAL: p.custo == null || isNaN(p.custo) ? '---' : p.custo,
            }
      );

    const headers = isAlpha
      ? ['ETIQUETA', 'DESCRICAO', 'EAN', 'SALDO', 'CONTROLADO', 'CUSTO_GERENCIAL']
      : ['EMPRESA', 'PRODUTO', 'DESCRICAO', 'EAN', 'SALDO', 'CONTROLADO', 'CUSTO_GERENCIAL'];
    const ws = XLSX.utils.json_to_sheet(dados, { header: headers });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Cadastro');
    const excelBuffer: any = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const data: Blob = new Blob([excelBuffer], { type: 'application/octet-stream' });
    FileSaver.saveAs(data, 'cadastro.xlsx');
  }

  exportarContagemPdf() {
    const doc = new jsPDF({ orientation: 'landscape' });

    doc.setFontSize(16);
    doc.setFont('helvetica', 'bold');
    doc.text('CONTAGEM DE PRODUTOS', doc.internal.pageSize.getWidth() / 2, 15, { align: 'center' });

    autoTable(doc, {
      startY: 25,
      head: [['Código', 'EAN', 'Descrição', 'Seção', 'Quantidade Escaneada']],
      body: this.contagemDetalhada()
        .slice()
        .sort((a, b) => a.descricao.localeCompare(b.descricao))
        .map((c) => [c.codigo, c.ean, c.descricao, c.secao, c.qtdeEscaneada]),
    });

    const pageHeight = doc.internal.pageSize.getHeight();
    const pageWidth = doc.internal.pageSize.getWidth();
    doc.setFontSize(10);
    doc.setTextColor(0, 0, 0);
    doc.setFont('helvetica', 'normal');
    doc.text(this.footerText, pageWidth / 2, pageHeight - 10, { align: 'center' });

    doc.save('contagem.pdf');
  }

  exportarContagemExcel() {
    const dados = this.contagemDetalhada()
      .slice()
      .sort((a, b) => a.descricao.localeCompare(b.descricao))
      .map((c) => ({
        codigo: c.codigo,
        ean: c.ean,
        descricao: c.descricao,
        secao: c.secao,
        qtdeEscaneada: c.qtdeEscaneada,
      }));
    const ws = XLSX.utils.json_to_sheet(dados, {
      header: ['codigo', 'ean', 'descricao', 'secao', 'qtdeEscaneada'],
    });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Contagem');
    const excelBuffer: any = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const data: Blob = new Blob([excelBuffer], { type: 'application/octet-stream' });
    FileSaver.saveAs(data, 'contagem.xlsx');
  }

  exportarContagemTxt() {
    try {
      const isAlpha = this.isAlpha7();
      const separator = ',';
      const agora = new Date();

      const linhas = this.contagemDetalhada()
        .slice()
        .sort((a, b) => String(a.descricao || '').localeCompare(String(b.descricao || '')))
        .map((c) => {
          const etiqueta = String(c.codigo ?? '').trim();
          const eanPrincipal = (this.splitEans(c.ean)[0] ?? '').toString();
          const identificador = isAlpha ? (etiqueta || eanPrincipal) : eanPrincipal;
          const qtde = Number(c.qtdeEscaneada ?? 0);

          if (isAlpha) {
            const data = this.formatarData(agora);
            const hora = this.formatarHora(agora);
            return [data, hora, identificador, qtde].join(separator);
          }

          return `${identificador}${separator}${qtde}`;
        });

      const conteudo = linhas.length ? linhas.join('\n') + '\n' : '';
      const blob = new Blob([conteudo], { type: 'text/plain;charset=utf-8' });
      FileSaver.saveAs(blob, 'contagem.txt');
    } catch (error) {
      console.error('Erro ao exportar TXT da contagem:', error);
      this.messageService.add({
        severity: 'error',
        summary: 'Erro ao exportar TXT',
        detail: 'N�o foi poss�vel gerar o arquivo contagem.txt',
      });
    }
  }

  private formatarData(data: Date): string {
    const ano = data.getFullYear();
    const mes = String(data.getMonth() + 1).padStart(2, '0');
    const dia = String(data.getDate()).padStart(2, '0');
    return `${ano}${mes}${dia}`;
  }

  private formatarHora(data: Date): string {
    const horas = String(data.getHours()).padStart(2, '0');
    const minutos = String(data.getMinutes()).padStart(2, '0');
    const segundos = String(data.getSeconds()).padStart(2, '0');
    return `${horas}${minutos}${segundos}`;
  }

  addProdutoContagem(rawEan: string, rawQtde: string) {
    try {
      const normalizeCodigo = (v: any) => String(v ?? '').trim().replace(/^0+/, '');
      const normalizeEan = (v: any) => String(v ?? '').trim().replace(/\D/g, '');

      const eanKey = normalizeEan(rawEan);
      const qtde = Number(rawQtde);

      if (!eanKey) {
        this.messageService.add({ severity: 'warn', summary: 'EAN inválido', detail: 'Informe um EAN válido.' });
        return;
      }
      if (!qtde || isNaN(qtde) || qtde <= 0) {
        this.messageService.add({ severity: 'warn', summary: 'Quantidade inválida', detail: 'Informe uma quantidade maior que zero.' });
        return;
      }

      // Mapear cadastro por EAN
      const cadastro = this.cadastro();
      const eanMap = new Map<string, any>();
      (cadastro || []).forEach((p: any) => {
        this.splitEans(p?.ean).forEach((e) => {
          const k = normalizeEan(e);
          if (k) eanMap.set(k, p);
        });
      });

      const prod = eanMap.get(eanKey);

      const lista = [...this.contagemDetalhada()];
      const targetKey = prod ? normalizeCodigo(prod.codigo) : eanKey;
      const idx = lista.findIndex(
        (i) => normalizeCodigo(i?.codigo) === targetKey || normalizeEan(i?.ean) === eanKey
      );

      if (idx >= 0) {
        // Atualiza item existente somando a quantidade
        const item = { ...lista[idx] };
        item.qtdeEscaneada = Number(item.qtdeEscaneada || 0) + qtde;
        lista[idx] = item;
      } else {
        // Insere novo item
        if (prod) {
          lista.push({
            codigo: prod.codigo,
            ean: prod.ean,
            descricao: prod.descricao || '',
            fabricante: prod.fabricante || '',
            secao: prod.secao || '',
            qtdeEscaneada: qtde,
          });
        } else {
          // Sem cadastro — ainda assim adiciona à contagem
          lista.push({
            codigo: eanKey,
            ean: eanKey,
            descricao: '',
            fabricante: '',
            secao: '',
            qtdeEscaneada: qtde,
          });
        }
      }

      this.contagemDetalhada.set(lista);
      this.gerarDivergencias();
      // Persistir contagem ao adicionar item
      this.indexedDBService
        .setItem('contagemDetalhada', this.contagemDetalhada())
        .catch(() => {});

      this.messageService.add({
        severity: prod ? 'success' : 'warn',
        summary: prod ? 'Produto adicionado' : 'Adicionado sem cadastro',
        detail: prod
          ? `${prod.descricao || 'Produto'} (Código ${prod.codigo})  +${qtde} unidades.`
          : `EAN ${eanKey} – +${qtde} unidades (produto não encontrado no cadastro).`,
      });
    } catch (error) {
      console.error('Erro ao adicionar produto na contagem:', error);
      this.messageService.add({ severity: 'error', summary: 'Erro', detail: 'Não foi possível adicionar o produto na contagem.' });
    }
  }

  removerProdutoContagem(item: any) {
    try {
      const normalizeCodigo = (v: any) => String(v ?? '').trim().replace(/^0+/, '');
      const normalizeEan = (v: any) => String(v ?? '').trim().replace(/\D/g, '');

      const cod = normalizeCodigo(item?.codigo);
      const ean = normalizeEan((this.splitEans(item?.ean)[0] ?? ''));

      const lista = this.contagemDetalhada().filter((i) => {
        const icod = normalizeCodigo(i?.codigo);
        const iean = normalizeEan((this.splitEans(i?.ean)[0] ?? ''));
        return !(icod === cod || iean === ean);
      });

      this.contagemDetalhada.set(lista);
      this.gerarDivergencias();
      // Persistir contagem ao remover item
      this.indexedDBService
        .setItem('contagemDetalhada', this.contagemDetalhada())
        .catch(() => {});

      this.messageService.add({
        severity: 'success',
        summary: 'Produto removido',
        detail: `${item?.descricao || 'Produto'} (Código ${item?.codigo ?? '-'}) removido da contagem.`,
      });
    } catch (error) {
      console.error('Erro ao remover produto da contagem:', error);
      this.messageService.add({ severity: 'error', summary: 'Erro', detail: 'Não foi possível remover o produto.' });
    }
  }

  exportarNaoContadosPdf() {
    const doc = new jsPDF({ orientation: 'landscape' });

    doc.setFontSize(16);
    doc.setFont('helvetica', 'bold');
    doc.text('PRODUTOS NÃO CONTADOS COM ESTOQUE', doc.internal.pageSize.getWidth() / 2, 15, { align: 'center' });

    autoTable(doc, {
      startY: 25,
      head: [['Código', 'EAN', 'Descrição', 'Seção', 'Quantidade']],
      body: this.naoInventariados()
        .slice()
        .sort((a, b) => a.descricao.localeCompare(b.descricao))
        .map((p) => [p.codigo, p.ean, p.descricao, p.secao, p.qtde]),
    });

    const pageHeight = doc.internal.pageSize.getHeight();
    const pageWidth = doc.internal.pageSize.getWidth();
    doc.setFontSize(10);
    doc.setTextColor(0, 0, 0);
    doc.setFont('helvetica', 'normal');
    doc.text(this.footerText, pageWidth / 2, pageHeight - 10, { align: 'center' });

    doc.save('nao-contados.pdf');
  }

  exportarNaoContadosExcel() {
    const dados = this.naoInventariados()
      .slice()
      .sort((a, b) => a.descricao.localeCompare(b.descricao))
      .map((p) => ({
        codigo: p.codigo,
        ean: p.ean,
        descricao: p.descricao,
        secao: p.secao,
        qtde: p.qtde,
      }));
    const ws = XLSX.utils.json_to_sheet(dados, {
      header: ['codigo', 'ean', 'descricao', 'secao', 'qtde'],
    });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Não Contados');
    const excelBuffer: any = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const data: Blob = new Blob([excelBuffer], { type: 'application/octet-stream' });
    FileSaver.saveAs(data, 'nao-contados.xlsx');
  }

}
