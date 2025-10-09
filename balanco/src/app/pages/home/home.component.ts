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
import { CheckboxModule } from 'primeng/checkbox';
import { bootstrapApplication, BrowserModule } from '@angular/platform-browser';
import { FormsModule } from '@angular/forms';
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

@Component({
  selector: 'app-home',
  templateUrl: './home.component.html',
  styleUrls: ['./home.component.scss'],
  standalone: true,
  imports: [
    HttpClientModule,
    CommonModule,
    DialogModule,
    FileUploadModule,
    ButtonModule,
    ToolbarModule,
    TableModule,
    InputTextModule,
    TabViewModule,
    ToastModule,
    CheckboxModule,
    FormsModule,
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
  divergencias = signal<any[]>([]);

  contagemDetalhada = signal<any[]>([]);
  naoInventariados = signal<any[]>([]);
  mostrarControladosContagem = signal(false);
  mostrarNaoCadastradosContagem = signal(false);
  mostrarControladosDivergencias = signal(false);
  mostrarNaoCadastradosDivergencias = signal(false);
  contagemFiltrada = computed(() => {
    const normalizar = (valor: any) =>
      String(valor ?? '')
        .trim()
        .toUpperCase();
    const isControlado = (item: any) => normalizar(item?.controlado).startsWith('S');
    const listaBase = this.contagemDetalhada();
    let lista = listaBase;
    if (this.mostrarControladosContagem()) {
      lista = lista.filter((item) => isControlado(item));
    }
    if (this.mostrarNaoCadastradosContagem()) {
      lista = lista.filter((item) => !!item?.semCadastro);
    }
    return lista;
  });
  divergenciasFiltradas = computed(() => {
    const normalizar = (valor: any) =>
      String(valor ?? '')
        .trim()
        .toUpperCase();
    const isControlado = (item: any) => normalizar(item?.controlado).startsWith('S');
    const listaBase = this.divergencias();
    let lista = listaBase;
    if (this.mostrarControladosDivergencias()) {
      lista = lista.filter((item) => isControlado(item));
    }
    if (this.mostrarNaoCadastradosDivergencias()) {
      lista = lista.filter(
        (item) => !!item?.semCadastro || normalizar(item?.tipo) === 'PRODUTO NÃO CADASTRADO'
      );
    }
    return lista;
  });
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
    'Razão Social:Drogarias Campeã  CNPJ: 21.812.204/0010-98 Endereço: Rua Santa Mônica, 480, Parque Industrial San José, Cotia-SP, CEP: 06715-865';

  private indexedDBService = new IndexedDBService('GestaoBalancoDB', 'CadastroProdutos');

  constructor(private messageService: MessageService) { }
  async ngOnInit() {
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

  onSelectedFiles(event: any) {
    console.log('Arquivos selecionados:', event?.files ?? event);
  }

  choose(event: Event, callback: Function) {
    event?.preventDefault();
    if (callback && typeof callback === 'function') {
      callback();
    }
  }

  uploadEvent(callback: Function) {
    if (callback && typeof callback === 'function') {
      callback();
    }
  }

  formatSize(bytes: number): string {
    if (!bytes) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return `${(bytes / Math.pow(k, i)).toFixed(2)} ${sizes[i]}`;
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

  gerarDivergencias() {
    const contagem = this.contagemDetalhada();
    const cadastro = this.cadastro();


    const cadastroMap = new Map(cadastro.map(p => [this.normalizeCodigo(p.codigo), p]));
    const contagemMap = new Map(contagem.map(c => [this.normalizeCodigo(c.codigo), c]));


    const divergencias: any[] = [];
    const naoContados: any[] = []; // NOVO


    contagem.forEach(c => {
      const cod = this.normalizeCodigo(c.codigo);
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
      const cod = this.normalizeCodigo(cad.codigo);
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

      this.lerCadastroProcfit(file)
        .then((cadastroProdutos) =>
          this.persistirCadastroProdutos(cadastroProdutos)
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
            const explicitKeys = ['EAN 1', 'CODIGO_EAN', 'CODIGO DE BARRAS', 'COD_BARRAS'];
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

  private persistirCadastroProdutos(cadastroProdutos: any[]) {
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
          detail: `${normalizados.length} produtos recuperados do IndexedDB (Procfit).`,
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
              ? `${total} produtos carregados (Procfit).`
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

  private normalizeCodigo(value: any): string {
    if (value === undefined || value === null) return '';
    const base = String(value)
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .toUpperCase()
      .trim();
    const alfanumerico = base.replace(/[^0-9A-Z]/g, '');
    const semZeros = alfanumerico.replace(/^0+/, '');
    return semZeros || alfanumerico || base;
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
      .split(/[;,\s\r\n\t|\/\\]+/)
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

      const normalizarCodigo = (v: any) => this.normalizeCodigo(v);

      // Acumulador por código
      const contagemMap = new Map<string, any>();

      this.indexedDBService.getItem('cadastroProdutos').then((cadastroProdutos) => {
        const cadastroFonte =
          Array.isArray(this.cadastro()) && this.cadastro().length
            ? this.cadastro()
            : (cadastroProdutos || []);

        // Indexar por código normalizado e por cada EAN individual
        const codigoMap = new Map<string, any>();
        cadastroFonte.forEach((p: any) => {
          const key = normalizarCodigo(p?.codigo);
          if (key) codigoMap.set(key, p);
        });
        const limparValor = (valor: any) =>
          String(valor ?? '')
            .trim()
            .replace(/^["']+|["']+$/g, '');

        const gerarChaves = (valor: any) => {
          const bruto = limparValor(valor);
          const semAcento = bruto
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .toUpperCase();
          const alfanumerico = semAcento.replace(/[^0-9A-Z]/g, '');
          const digitos = bruto.replace(/\D/g, '');
          const digitosSemZero = digitos.replace(/^0+/, '');
          return {
            bruto,
            semAcento,
            alfanumerico,
            digitos,
            digitosSemZero,
          };
        };
        const adicionarChaves = (map: Map<string, any>, produto: any, chaves: (string | undefined)[]) => {
          chaves
            .map((c) => (c ?? '').trim())
            .filter((c) => !!c)
            .forEach((chave) => map.set(chave, produto));
        };

        const eanMap = new Map<string, any>();
        cadastroFonte.forEach((p: any) => {
          const codigoChaves = gerarChaves(p?.codigo);
          adicionarChaves(codigoMap, p, [
            codigoChaves.alfanumerico,
            codigoChaves.semAcento,
            codigoChaves.bruto,
          ]);

          this.splitEans(p.ean).forEach((e) => {
            const chaves = gerarChaves(e);
            adicionarChaves(eanMap, p, [
              chaves.digitosSemZero,
              chaves.digitos,
              chaves.alfanumerico,
              chaves.semAcento,
              chaves.bruto,
            ]);
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
                  const [eanValue, qtdeValue, secaoValue, coletorValue, inventariadorValue] = l.split(/[|,;]/);
                  const raw = limparValor(eanValue);
                  const chavesEntrada = gerarChaves(raw);
                  const keyCodigo = normalizarCodigo(raw);
                  const secaoLinha = limparValor(secaoValue);
                  const coletorLinha = limparValor(coletorValue);
                  const inventariadorLinha = limparValor(inventariadorValue);

                  const prod =
                    codigoMap.get(keyCodigo) ??
                    codigoMap.get(chavesEntrada.alfanumerico) ??
                    codigoMap.get(chavesEntrada.semAcento) ??
                    codigoMap.get(chavesEntrada.bruto) ??
                    eanMap.get(chavesEntrada.digitosSemZero) ??
                    eanMap.get(chavesEntrada.digitos) ??
                    eanMap.get(chavesEntrada.alfanumerico) ??
                    eanMap.get(chavesEntrada.semAcento) ??
                    eanMap.get(chavesEntrada.bruto);
                  const semCadastro = !prod;
                  const contKeyBase = prod ? normalizarCodigo(prod.codigo) : keyCodigo;
                  const contKey =
                    contKeyBase ||
                    chavesEntrada.digitosSemZero ||
                    chavesEntrada.digitos ||
                    chavesEntrada.alfanumerico ||
                    chavesEntrada.semAcento ||
                    chavesEntrada.bruto;

                  let atual = contagemMap.get(contKey);
                  if (!atual) {
                    atual = {
                      codigo: prod?.codigo ?? raw,
                      ean: prod?.ean ?? raw,
                      descricao: semCadastro
                        ? 'PRODUTO NÃO ENCONTRADO NO CADASTRO'
                        : (prod?.descricao ?? ''),
                      fabricante: semCadastro ? '---' : (prod?.fabricante ?? '---'),
                      secao: semCadastro ? (secaoLinha || '---') : (secaoLinha || prod?.secao || ''),
                      qtdeEscaneada: 0,
                      coletor: coletorLinha,
                      inventariador: inventariadorLinha,
                      semCadastro,
                    };
                  } else {
                    if (coletorLinha) {
                      atual.coletor = coletorLinha;
                    }
                    if (inventariadorLinha) {
                      atual.inventariador = inventariadorLinha;
                    }
                  }

                  if (prod) {
                    atual.semCadastro = false;
                    atual.codigo = prod.codigo ?? atual.codigo;
                    atual.ean = prod.ean ?? atual.ean;
                    atual.descricao = prod.descricao ?? atual.descricao ?? '';
                    atual.fabricante = prod.fabricante ?? atual.fabricante ?? '---';
                    if (secaoLinha) {
                      atual.secao = secaoLinha;
                    } else if (!atual.secao && prod.secao) {
                      atual.secao = prod.secao;
                    }
                    atual.controlado = prod.controlado ?? atual.controlado ?? '';
                  } else {
                    atual.semCadastro = true;
                    atual.descricao = 'PRODUTO NÃO ENCONTRADO NO CADASTRO';
                    atual.fabricante = '---';
                    if (secaoLinha) {
                      atual.secao = secaoLinha;
                    } else if (!atual.secao) {
                      atual.secao = '---';
                    }
                    atual.controlado = 'N';
                  }

                  atual.qtdeEscaneada += Number(qtdeValue) || 0;
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

    autoTable(doc, {
      startY: 25,
      head: [
        [
          'Código',
          'EAN 1',
          'Descrição',
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
          d.codigo ?? '---',
          this.splitEans(d.ean)[0] ?? '---',
          d.descricao ?? '---',
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

    const finalY = (doc as any).lastAutoTable?.finalY ?? 25;

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
        coletor: c.coletor ?? '',
        inventariador: c.inventariador ?? '',
      }));
    const ws = XLSX.utils.json_to_sheet(dados, {
      header: ['codigo', 'EAN 1', 'descricao', 'secao', 'qtdeEscaneada', 'coletor', 'inventariador'],
    });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Contagem');
    const excelBuffer: any = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const data: Blob = new Blob([excelBuffer], { type: 'application/octet-stream' });
    FileSaver.saveAs(data, 'contagem.xlsx');
  }

  exportarContagemTxt() {
    try {
      const separator = ',';

      const linhas = this.contagemDetalhada()
        .slice()
        .sort((a, b) => String(a.descricao || '').localeCompare(String(b.descricao || '')))
        .map((c) => {
          const eanPrincipal = (this.splitEans(c.ean)[0] ?? '').toString();
          const qtde = Number(c.qtdeEscaneada ?? 0);
          return `${eanPrincipal}${separator}${qtde}`;
        });

      const conteudo = linhas.length ? linhas.join('\n') + '\n' : '';
      const blob = new Blob([conteudo], { type: 'text/plain;charset=utf-8' });
      FileSaver.saveAs(blob, 'contagem.txt');
    } catch (error) {
      console.error('Erro ao exportar TXT da contagem:', error);
      this.messageService.add({
        severity: 'error',
        summary: 'Erro ao exportar TXT',
        detail: 'Não foi possível gerar o arquivo contagem.txt',
      });
    }
  }

  addProdutoContagem(rawEan: string, rawQtde: string) {
    try {
      const normalizeCodigo = (v: any) => this.normalizeCodigo(v);
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
      const targetKey = (prod ? normalizeCodigo(prod.codigo) : '') || eanKey;
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
            coletor: '',
            inventariador: '',
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
            coletor: '',
            inventariador: '',
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
      const normalizeCodigo = (v: any) => this.normalizeCodigo(v);
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

  exportarCadastroPdf() {
    const doc = new jsPDF({ orientation: 'landscape' });

    doc.setFontSize(16);
    doc.setFont('helvetica', 'bold');
    doc.text('CADASTRO DE PRODUTOS', doc.internal.pageSize.getWidth() / 2, 15, { align: 'center' });

    autoTable(doc, {
      startY: 25,
      head: [['Produto', 'EAN', 'Descrição', 'Saldo', 'Controlado', 'Custo Gerencial', 'Empresa', 'Seção']],
      body: this.cadastro()
        .slice()
        .sort((a, b) => a.descricao.localeCompare(b.descricao))
        .map((p) => [
          p.codigo ?? '---',
          this.splitEans(p.ean)[0] ?? '---',
          p.descricao ?? '---',
          p.qtde ?? '---',
          p.controlado ?? '---',
          p.custo == null || isNaN(p.custo) ? '---' : Number(p.custo).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }),
          p.fabricante ?? '---',
          p.secao ?? '---',
        ]),
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
    const dados = this.cadastro()
      .slice()
      .sort((a, b) => a.descricao.localeCompare(b.descricao))
      .map((p) => ({
        PRODUTO: p.codigo,
        'EAN 1': this.splitEans(p.ean)[0] ?? '',
        DESCRICAO: p.descricao,
        SALDO: p.qtde,
        CONTROLADO: p.controlado,
        'CUSTO GERENCIAL': p.custo == null || isNaN(p.custo) ? '---' : p.custo,
        EMPRESA: p.fabricante ?? '',
        SECAO: p.secao ?? '',
      }));
    const ws = XLSX.utils.json_to_sheet(dados, {
      header: ['PRODUTO', 'EAN 1', 'DESCRICAO', 'SALDO', 'CONTROLADO', 'CUSTO GERENCIAL', 'EMPRESA', 'SECAO'],
    });
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
      head: [['Produto', 'EAN', 'Descrição', 'Seção', 'Quantidade Escaneada', 'Coletor', 'Inventariador']],
      body: this.contagemDetalhada()
        .slice()
        .sort((a, b) => a.descricao.localeCompare(b.descricao))
        .map((c) => [
          c.codigo ?? '---',
          c.ean ?? '---',
          c.descricao ?? '---',
          c.secao ?? '---',
          c.qtdeEscaneada ?? '---',
          c.coletor ?? '',
          c.inventariador ?? '',
        ]),
    });

    const pageHeight = doc.internal.pageSize.getHeight();
    const pageWidth = doc.internal.pageSize.getWidth();
    doc.setFontSize(10);
    doc.setTextColor(0, 0, 0);
    doc.setFont('helvetica', 'normal');
    doc.text(this.footerText, pageWidth / 2, pageHeight - 10, { align: 'center' });

    doc.save('contagem.pdf');
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
        ean1: this.splitEans(d.ean)[0] ?? '---',
        descricao: d.descricao ?? '---',
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
        'ean1',
        'descricao',
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
    FileSaver.saveAs(data, 'divergencias.xlsx');
  }

  get totalPositivo(): number {
    return this.divergenciasFiltradas()
      .filter((d) => d.valorDiferenca != null && d.valorDiferenca > 0)
      .reduce((sum, d) => sum + d.valorDiferenca, 0);
  }

  get totalNegativo(): number {
    return this.divergenciasFiltradas()
      .filter((d) => d.valorDiferenca != null && d.valorDiferenca < 0)
      .reduce((sum, d) => sum + d.valorDiferenca, 0);
  }

  get diferencaBalanco(): number {
    return this.totalPositivo + this.totalNegativo;
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




