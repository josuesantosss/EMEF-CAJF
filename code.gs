/**
 * Sistema de Gestão Escolar - Google Apps Script
 * Integração com planilha do Google Sheets
 * Versão 2.0 - Atualizada
 */

// Configuração da planilha
const SPREADSHEET_ID = '1vfgwXNrZuQ8KBcjCuJB_5RKUQ3imAYbAHsMX1yG3It0';

// ============================================
// FUNÇÕES DE UTILIDADE
// ============================================

function gerarID() {
    return 'ID_' + new Date().getTime() + '_' + Math.floor(Math.random() * 10000);
}

function formatarData(data) {
    if (!data) return '';
    if (typeof data === 'string') return data;
    const d = new Date(data);
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function formatarDataHora(data) {
    if (!data) return '';
    const d = new Date(data);
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function getOrCreateSheet(nomeAba, cabecalho) {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(nomeAba);
    
    if (!sheet) {
        sheet = ss.insertSheet(nomeAba);
        if (cabecalho && cabecalho.length > 0) {
            sheet.getRange(1, 1, 1, cabecalho.length).setValues([cabecalho]);
            sheet.getRange(1, 1, 1, cabecalho.length).setFontWeight('bold');
            sheet.getRange(1, 1, 1, cabecalho.length).setBackground('#f0f0f0');
        }
    }
    
    return sheet;
}

function adicionarLinha(nomeAba, dados, cabecalho) {
    try {
        const sheet = getOrCreateSheet(nomeAba, cabecalho);
        const ultimaLinha = sheet.getLastRow();
        const novaLinha = ultimaLinha + 1;
        
        const linhaDados = cabecalho.map(col => {
            return dados[col] !== undefined ? dados[col] : '';
        });
        
        sheet.getRange(novaLinha, 1, 1, linhaDados.length).setValues([linhaDados]);
        
        return { success: true, linha: novaLinha };
    } catch (error) {
        console.error('Erro ao adicionar linha:', error);
        return { success: false, error: error.message };
    }
}

function buscarDados(nomeAba, colunaBusca, valorBusca, cabecalho) {
    try {
        const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(nomeAba);
        if (!sheet) return [];
        
        const dados = sheet.getDataRange().getValues();
        if (dados.length < 2) return [];
        
        const colIndex = cabecalho.indexOf(colunaBusca);
        if (colIndex === -1) return [];
        
        const resultados = [];
        for (let i = 1; i < dados.length; i++) {
            if (String(dados[i][colIndex]).trim() === String(valorBusca).trim()) {
                const registro = {};
                cabecalho.forEach((col, j) => {
                    registro[col] = dados[i][j] !== undefined ? dados[i][j] : '';
                });
                resultados.push(registro);
            }
        }
        
        return resultados;
    } catch (error) {
        console.error('Erro ao buscar dados:', error);
        return [];
    }
}

function atualizarLinha(nomeAba, colunaBusca, valorBusca, dadosAtualizados, cabecalho) {
    try {
        const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(nomeAba);
        if (!sheet) return { success: false, error: 'Aba não encontrada' };
        
        const dados = sheet.getDataRange().getValues();
        if (dados.length < 2) return { success: false, error: 'Nenhum dado encontrado' };
        
        const colIndex = cabecalho.indexOf(colunaBusca);
        if (colIndex === -1) return { success: false, error: 'Coluna de busca não encontrada' };
        
        let linhaEncontrada = -1;
        for (let i = 1; i < dados.length; i++) {
            if (String(dados[i][colIndex]).trim() === String(valorBusca).trim()) {
                linhaEncontrada = i + 1;
                break;
            }
        }
        
        if (linhaEncontrada === -1) {
            return { success: false, error: 'Registro não encontrado' };
        }
        
        const linhaAtual = dados[linhaEncontrada - 1];
        cabecalho.forEach((col, index) => {
            if (dadosAtualizados[col] !== undefined) {
                linhaAtual[index] = dadosAtualizados[col];
            }
        });
        
        sheet.getRange(linhaEncontrada, 1, 1, linhaAtual.length).setValues([linhaAtual]);
        
        return { success: true, linha: linhaEncontrada };
    } catch (error) {
        console.error('Erro ao atualizar linha:', error);
        return { success: false, error: error.message };
    }
}

// ============================================
// FUNÇÃO PARA NORMALIZAR NOME DA ABA DE PRESENÇA
// ============================================

function getNomeAbaPresenca(turma) {
    // Mapeamento de turmas para nomes das abas
    const mapaAbas = {
        '1A': 'PRESENCAS_1A',
        '1B': 'PRESENCAS_1B',
        '1C': 'PRESENCAS_1C',
        '1D': 'PRESENCAS_1D',
        '1E': 'PRESENCAS_1E',
        '2A': 'PRESENCAS_2A',
        '2B': 'PRESENCAS_2B',
        '2C': 'PRESENCAS_2C',
        '2D': 'PRESENCAS_2D',
        '2E': 'PRESENCAS_2E',
        '3A': 'PRESENCAS_3A',
        '3B': 'PRESENCAS_3B',
        '3C': 'PRESENCAS_3C',
        '3D': 'PRESENCAS_3D',
        '3E': 'PRESENCAS_3E',
        '4A': 'PRESENCAS_4A',
        '4B': 'PRESENCAS_4B',
        '4C': 'PRESENCAS_4C',
        '4D': 'PRESENCAS_4D',
        '4E': 'PRESENCAS_4E',
        '5A': 'PRESENCAS_5A',
        '5B': 'PRESENCAS_5B',
        '5C': 'PRESENCAS_5C',
        '5D': 'PRESENCAS_5D',
        '5E': 'PRESENCAS_5E'
    };
    
    // Normaliza a turma para maiúsculas e remove espaços
    const turmaNormalizada = String(turma).toUpperCase().trim().replace(/\s/g, '');
    
    // Tenta encontrar no mapa
    if (mapaAbas[turmaNormalizada]) {
        return mapaAbas[turmaNormalizada];
    }
    
    // Se não encontrar, tenta extrair o número e a letra
    const match = turmaNormalizada.match(/(\d+)([A-Z])/);
    if (match) {
        const numero = match[1];
        const letra = match[2];
        return `PRESENCA_${numero}${letra}`;
    }
    
    // Fallback: usa o nome da turma com prefixo
    return `PRESENCA_${turmaNormalizada}`;
}

function getNomeAbaNotas(turma) {
    // Normaliza a turma para maiúsculas e remove espaços
    const turmaNormalizada = String(turma).toUpperCase().trim().replace(/\s/g, '');
    
    // Retorna o nome da aba de notas
    return `NOTAS_BIMESTRAIS_${turmaNormalizada}`;
}

// ============================================
// FUNÇÕES DE ATRIBUIÇÃO
// ============================================

function atribuirTurma(professorId, professorNome, tipo, turma, dataAtribuicao) {
    try {
        const cabecalho = ['ID_Professor', 'Nome_Professor', 'Tipo', 'Turma', 'Data_Atribuicao', 'Ativo'];
        
        // Verifica se já existe atribuição ativa para esta turma
        const atribuicoesExistentes = buscarDados('ATRIBUICOES', 'Turma', turma, cabecalho);
        const atribuicaoAtiva = atribuicoesExistentes.find(a => 
            String(a.Ativo).toUpperCase() === 'SIM'
        );
        
        if (atribuicaoAtiva) {
            return { 
                success: false, 
                error: `A turma ${turma} já está atribuída ao professor ${atribuicaoAtiva.Nome_Professor}` 
            };
        }
        
        const dados = {
            ID_Professor: professorId,
            Nome_Professor: professorNome,
            Tipo: tipo || 'Regente',
            Turma: turma,
            Data_Atribuicao: dataAtribuicao || formatarData(new Date()),
            Ativo: 'SIM'
        };
        
        return adicionarLinha('ATRIBUICOES', dados, cabecalho);
    } catch (error) {
        console.error('Erro ao atribuir turma:', error);
        return { success: false, error: error.message };
    }
}

function removerTurma(professorId, turma) {
    try {
        const cabecalho = ['ID_Professor', 'Nome_Professor', 'Tipo', 'Turma', 'Data_Atribuicao', 'Ativo'];
        
        const atribuicoes = buscarDados('ATRIBUICOES', 'ID_Professor', professorId, cabecalho);
        const atribuicao = atribuicoes.find(a => 
            String(a.Turma).trim() === String(turma).trim() && 
            String(a.Ativo).toUpperCase() === 'SIM'
        );
        
        if (!atribuicao) {
            return { success: false, error: 'Atribuição não encontrada' };
        }
        
        return atualizarLinha('ATRIBUICOES', 'Turma', turma, {
            Ativo: 'NÃO'
        }, cabecalho);
    } catch (error) {
        console.error('Erro ao remover turma:', error);
        return { success: false, error: error.message };
    }
}

// ============================================
// FUNÇÕES DE RELATÓRIOS
// ============================================

function solicitarRelatorio(turma, aluno, bimestre, professor) {
    try {
        const cabecalho = [
            'ID_Solicitacao', 
            'Data_Solicitacao', 
            'Turma', 
            'Aluno', 
            'Bimestre', 
            'Professor', 
            'Status', 
            'Data_Conclusao', 
            'Observacoes'
        ];
        
        const dados = {
            ID_Solicitacao: gerarID(),
            Data_Solicitacao: formatarData(new Date()),
            Turma: turma,
            Aluno: aluno,
            Bimestre: bimestre,
            Professor: professor || '',
            Status: 'Pendente',
            Data_Conclusao: '',
            Observacoes: ''
        };
        
        return adicionarLinha('RELATORIOS_SOLICITADOS', dados, cabecalho);
    } catch (error) {
        console.error('Erro ao solicitar relatório:', error);
        return { success: false, error: error.message };
    }
}

function enviarRelatorio(turma, aluno, bimestre, professor, notas, media, presenca, observacoes, recomendacao, conteudo) {
    try {
        const cabecalho = [
            'ID_Relatorio',
            'Data_Geracao',
            'Turma',
            'Aluno',
            'Bimestre',
            'Professor',
            'Notas',
            'Media',
            'Presenca',
            'Observacoes',
            'Recomendacao',
            'Conteudo',
            'Visto',
            'Data_Visto'
        ];
        
        // Verifica se já existe relatório para este aluno/bimestre/professor
        const existentes = buscarDados('RELATORIOS_GERADOS', 'Turma', turma, cabecalho);
        const jaExiste = existentes.some(r => 
            String(r.Aluno).trim() === String(aluno).trim() && 
            String(r.Bimestre).trim() === String(bimestre).trim() &&
            String(r.Professor).trim() === String(professor).trim()
        );
        
        if (jaExiste) {
            const dadosAtualizados = {
                Data_Geracao: formatarData(new Date()),
                Notas: notas || '',
                Media: media || '',
                Presenca: presenca || '',
                Observacoes: observacoes || '',
                Recomendacao: recomendacao || '',
                Conteudo: conteudo || '',
                Visto: '',
                Data_Visto: ''
            };
            
            // Atualiza o relatório existente
            const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('RELATORIOS_GERADOS');
            if (sheet) {
                const dados = sheet.getDataRange().getValues();
                for (let i = 1; i < dados.length; i++) {
                    if (String(dados[i][3]).trim() === String(aluno).trim() && 
                        String(dados[i][4]).trim() === String(bimestre).trim() &&
                        String(dados[i][5]).trim() === String(professor).trim()) {
                        
                        cabecalho.forEach((col, index) => {
                            if (dadosAtualizados[col] !== undefined) {
                                sheet.getRange(i + 1, index + 1).setValue(dadosAtualizados[col]);
                            }
                        });
                        break;
                    }
                }
            }
            
            return { success: true, message: 'Relatório atualizado com sucesso' };
        }
        
        const dados = {
            ID_Relatorio: gerarID(),
            Data_Geracao: formatarData(new Date()),
            Turma: turma,
            Aluno: aluno,
            Bimestre: bimestre,
            Professor: professor || '',
            Notas: notas || '',
            Media: media || '',
            Presenca: presenca || '',
            Observacoes: observacoes || '',
            Recomendacao: recomendacao || '',
            Conteudo: conteudo || '',
            Visto: '',
            Data_Visto: ''
        };
        
        // Marca a solicitação como concluída
        const cabecalhoSolicitacao = [
            'ID_Solicitacao', 
            'Data_Solicitacao', 
            'Turma', 
            'Aluno', 
            'Bimestre', 
            'Professor', 
            'Status', 
            'Data_Conclusao', 
            'Observacoes'
        ];
        
        const solicitacoes = buscarDados('RELATORIOS_SOLICITADOS', 'Turma', turma, cabecalhoSolicitacao);
        const solicitacao = solicitacoes.find(s => 
            String(s.Aluno).trim() === String(aluno).trim() && 
            String(s.Bimestre).trim() === String(bimestre).trim() && 
            String(s.Status).toUpperCase() === 'PENDENTE'
        );
        
        if (solicitacao) {
            atualizarLinha('RELATORIOS_SOLICITADOS', 'ID_Solicitacao', solicitacao.ID_Solicitacao, {
                Status: 'Concluído',
                Data_Conclusao: formatarData(new Date())
            }, cabecalhoSolicitacao);
        }
        
        return adicionarLinha('RELATORIOS_GERADOS', dados, cabecalho);
    } catch (error) {
        console.error('Erro ao enviar relatório:', error);
        return { success: false, error: error.message };
    }
}

function marcarRelatorioComoVisto(idRelatorio, turma, aluno, bimestre, professor) {
    try {
        const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('RELATORIOS_GERADOS');
        
        if (!sheet) {
            return { success: false, error: 'Aba RELATORIOS_GERADOS não encontrada' };
        }
        
        const dados = sheet.getDataRange().getValues();
        const cabecalho = dados[0];
        
        // Procura a coluna "Visto" e "Data_Visto"
        let colVisto = -1;
        let colDataVisto = -1;
        
        cabecalho.forEach((col, index) => {
            if (String(col).trim() === 'Visto') colVisto = index;
            if (String(col).trim() === 'Data_Visto') colDataVisto = index;
        });
        
        // Se não existir a coluna "Visto", adiciona
        if (colVisto === -1) {
            colVisto = cabecalho.length;
            sheet.getRange(1, colVisto + 1).setValue('Visto');
            cabecalho.push('Visto');
        }
        
        // Se não existir a coluna "Data_Visto", adiciona
        if (colDataVisto === -1) {
            colDataVisto = colVisto + 1;
            sheet.getRange(1, colDataVisto + 1).setValue('Data_Visto');
            cabecalho.push('Data_Visto');
        }
        
        // Procura o relatório pelo ID ou combinação aluno/bimestre/professor
        let linhaEncontrada = -1;
        
        for (let i = 1; i < dados.length; i++) {
            const idAtual = String(dados[i][0] || '');
            const alunoAtual = String(dados[i][3] || '');
            const bimestreAtual = String(dados[i][4] || '');
            const professorAtual = String(dados[i][5] || '');
            
            if (idRelatorio && idAtual === idRelatorio) {
                linhaEncontrada = i + 1;
                break;
            } else if (aluno && bimestre && professor &&
                       alunoAtual === aluno && 
                       bimestreAtual === bimestre && 
                       professorAtual === professor) {
                linhaEncontrada = i + 1;
                break;
            }
        }
        
        if (linhaEncontrada === -1) {
            return { success: false, error: 'Relatório não encontrado' };
        }
        
        const dataHora = new Date();
        const dataFormatada = Utilities.formatDate(dataHora, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        const horaFormatada = Utilities.formatDate(dataHora, Session.getScriptTimeZone(), 'HH:mm:ss');
        
        // Marca como visto
        sheet.getRange(linhaEncontrada, colVisto + 1).setValue('SIM');
        sheet.getRange(linhaEncontrada, colDataVisto + 1).setValue(`${dataFormatada} ${horaFormatada}`);
        
        return { 
            success: true, 
            message: 'Relatório marcado como visto',
            dataVisto: `${dataFormatada} ${horaFormatada}`
        };
        
    } catch (error) {
        console.error('Erro ao marcar relatório como visto:', error);
        return { success: false, error: error.message };
    }
}

// ============================================
// FUNÇÕES DE NOTAS
// ============================================

const MAPA_MATERIAS = {
    'LP': { coluna: 4, prefixo: 'LP' },
    'MAT': { coluna: 5, prefixo: 'MAT' },
    'CIE': { coluna: 6, prefixo: 'CIE' },
    'HIST': { coluna: 7, prefixo: 'HIST' },
    'GEO': { coluna: 8, prefixo: 'GEO' },
    'ING': { coluna: 9, prefixo: 'ING' },
    'EDFIS': { coluna: 10, prefixo: 'EDFIS' },
    'ARTE': { coluna: 11, prefixo: 'ARTE' },
    'MUSICA': { coluna: 12, prefixo: 'MUSICA' }
};

function salvarNotaEstruturaCorreta(turma, aluno, disciplina, bimestre, nota) {
    try {
        const turmaFormatada = String(turma).toUpperCase().trim().replace(/\s/g, '');
        const nomeAba = getNomeAbaNotas(turmaFormatada);
        
        let sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(nomeAba);
        
        if (!sheet) {
            const cabecalho = [
                'ID_Registro', 'Turma', 'Nome_Aluno',
                'LP-1', 'MAT-1', 'CIE-1', 'HIST-1', 'GEO-1', 'ING-1', 'EDFIS-1', 'ARTE-1', 'MUSICA-1',
                'LP-2', 'MAT-2', 'CIE-2', 'HIST-2', 'GEO-2', 'ING-2', 'EDFIS-2', 'ARTE-2', 'MUSICA-2',
                'LP-3', 'MAT-3', 'CIE-3', 'HIST-3', 'GEO-3', 'ING-3', 'EDFIS-3', 'ARTE-3', 'MUSICA-3',
                'LP-4', 'MAT-4', 'CIE-4', 'HIST-4', 'GEO-4', 'ING-4', 'EDFIS-4', 'ARTE-4', 'MUSICA-4'
            ];
            sheet = getOrCreateSheet(nomeAba, cabecalho);
        }
        
        const dados = sheet.getDataRange().getValues();
        let linhaAluno = -1;
        
        // Procura o aluno na turma
        for (let i = 1; i < dados.length; i++) {
            const nomeAluno = String(dados[i][2] || '').trim();
            const turmaAluno = String(dados[i][1] || '').trim();
            
            if (nomeAluno === aluno && turmaAluno === turmaFormatada) {
                linhaAluno = i + 1;
                break;
            }
        }
        
        // Se o aluno não existe, adiciona nova linha
        if (linhaAluno === -1) {
            const novaLinha = dados.length + 1;
            const novaLinhaDados = new Array(39).fill(''); // 3 + 36 colunas
            novaLinhaDados[0] = gerarID();
            novaLinhaDados[1] = turmaFormatada;
            novaLinhaDados[2] = aluno;
            
            sheet.getRange(novaLinha, 1, 1, novaLinhaDados.length).setValues([novaLinhaDados]);
            linhaAluno = novaLinha;
        }
        
        // Determina a coluna baseada na disciplina e bimestre
        const infoMateria = MAPA_MATERIAS[disciplina.toUpperCase()];
        
        if (!infoMateria) {
            return { success: false, error: 'Disciplina não reconhecida: ' + disciplina };
        }
        
        // Coluna = coluna base da disciplina + (bimestre - 1) * 9
        const colunaNota = infoMateria.coluna + (bimestre - 1) * 9;
        
        sheet.getRange(linhaAluno, colunaNota).setValue(nota);
        
        return { success: true, linha: linhaAluno, coluna: colunaNota };
        
    } catch (error) {
        console.error('Erro ao salvar nota:', error);
        return { success: false, error: error.message };
    }
}

function salvarMultiplasNotas(notasArray) {
    try {
        let salvos = 0;
        let erros = [];
        
        notasArray.forEach(item => {
            const resultado = salvarNotaEstruturaCorreta(
                item.turma,
                item.aluno,
                item.disciplina || 'MAT',
                item.bimestre,
                item.nota
            );
            if (resultado.success) {
                salvos++;
            } else {
                erros.push(`${item.aluno}: ${resultado.error}`);
            }
        });
        
        return { success: salvos > 0, salvos, erros };
    } catch (error) {
        console.error('Erro ao salvar múltiplas notas:', error);
        return { success: false, error: error.message };
    }
}

// ============================================
// FUNÇÕES DE PRESENÇA - COM ABAS POR TURMA
// ============================================

function salvarPresenca(turma, data, alunosPresentes, alunosFaltantes, observacao) {
    try {
        // Converte para arrays se for string
        const presentes = Array.isArray(alunosPresentes) ? alunosPresentes : [];
        const faltantes = Array.isArray(alunosFaltantes) ? alunosFaltantes : [];
        
        // Determina o nome da aba com base na turma
        const nomeAba = getNomeAbaPresenca(turma);
        console.log(`📋 Salvando presença na aba: ${nomeAba}`);
        
        // Busca a aba específica da turma
        let sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(nomeAba);
        
        // Se não existir, cria com cabeçalho padrão
        if (!sheet) {
            const cabecalho = ['Nome_Aluno', 'Turma', 'Observacao', 'Data'];
            sheet = getOrCreateSheet(nomeAba, cabecalho);
        }
        
        // Obtém todos os dados
        const dados = sheet.getDataRange().getValues();
        
        // ============================================
        // 1. PROCURA OU CRIA A COLUNA DA DATA
        // ============================================
        let colunaData = -1;
        
        // Procura a data no cabeçalho (linha 0) a partir da coluna 4
        for (let j = 3; j < dados[0].length; j++) {
            if (String(dados[0][j]).trim() === String(data).trim()) {
                colunaData = j;
                break;
            }
        }
        
        // Se a data não existe, adiciona como nova coluna
        if (colunaData === -1) {
            const ultimaColuna = sheet.getLastColumn();
            colunaData = ultimaColuna; // índice 0-based
            sheet.getRange(1, ultimaColuna + 1).setValue(data);
            // Recarrega dados para atualizar o cabeçalho
            const novosDados = sheet.getDataRange().getValues();
            dados.length = 0;
            dados.push(...novosDados);
        }
        
        // ============================================
        // 2. BUSCA TODOS OS ALUNOS DA TURMA
        // ============================================
        const alunosDaTurma = [];
        for (let i = 1; i < dados.length; i++) {
            const nomeAluno = String(dados[i][0] || '').trim();
            const turmaLinha = String(dados[i][1] || '').trim();
            if (nomeAluno && (!turmaLinha || turmaLinha === turma)) {
                alunosDaTurma.push({ linha: i + 1, nome: nomeAluno });
            }
        }
        
        // ============================================
        // 3. MARCA PRESENÇA E FALTA PARA CADA ALUNO
        // ============================================
        let marcacoes = 0;
        
        // Combina todos os alunos relevantes (presentes + faltantes)
        const todosAlunos = [...new Set([...presentes, ...faltantes])];
        
        todosAlunos.forEach(aluno => {
            // Procura o aluno na lista da turma
            const alunoEncontrado = alunosDaTurma.find(a => 
                String(a.nome).trim() === String(aluno).trim()
            );
            
            if (alunoEncontrado) {
                // Aluno existe - marca na coluna da data
                const estaPresente = presentes.some(p => String(p).trim() === String(aluno).trim());
                const valor = estaPresente ? 'P' : 'F';
                
                sheet.getRange(alunoEncontrado.linha, colunaData + 1).setValue(valor);
                marcacoes++;
            } else {
                // Aluno não está na lista - adiciona nova linha
                const ultimaLinha = sheet.getLastRow();
                const novaLinha = ultimaLinha + 1;
                
                // Cria linha com o nome do aluno
                const numColunas = Math.max(dados[0].length, colunaData + 1, 4);
                const novaLinhaDados = new Array(numColunas).fill('');
                novaLinhaDados[0] = aluno;
                novaLinhaDados[1] = turma;
                novaLinhaDados[2] = observacao || '';
                
                // Marca P ou F na coluna da data
                const estaPresente = presentes.some(p => String(p).trim() === String(aluno).trim());
                if (colunaData < novaLinhaDados.length) {
                    novaLinhaDados[colunaData] = estaPresente ? 'P' : 'F';
                }
                
                sheet.getRange(novaLinha, 1, 1, novaLinhaDados.length).setValues([novaLinhaDados]);
                marcacoes++;
            }
        });
        
        return { success: true, marcacoes, aba: nomeAba };
        
    } catch (error) {
        console.error('Erro ao salvar presença:', error);
        return { success: false, error: error.message };
    }
}

// ============================================
// FUNÇÕES DE CONTEÚDOS E OBSERVAÇÕES
// ============================================

function salvarConteudo(turma, data, disciplina, conteudo, objetivos) {
    try {
        const cabecalho = ['Turma', 'Data', 'Disciplina', 'Conteudo', 'Objetivos'];
        
        const dados = {
            Turma: turma,
            Data: data || formatarData(new Date()),
            Disciplina: disciplina || '',
            Conteudo: conteudo || '',
            Objetivos: objetivos || ''
        };
        
        return adicionarLinha('CONTEUDOS', dados, cabecalho);
    } catch (error) {
        console.error('Erro ao salvar conteúdo:', error);
        return { success: false, error: error.message };
    }
}

function salvarObservacao(turma, data, aluno, observacao, tipo) {
    try {
        const cabecalho = ['Turma', 'Data', 'Aluno', 'Observacao', 'Tipo'];
        
        const dados = {
            Turma: turma,
            Data: data || formatarData(new Date()),
            Aluno: aluno || '',
            Observacao: observacao || '',
            Tipo: tipo || 'Neutra'
        };
        
        return adicionarLinha('OBSERVACOES', dados, cabecalho);
    } catch (error) {
        console.error('Erro ao salvar observação:', error);
        return { success: false, error: error.message };
    }
}

// ============================================
// FUNÇÕES PARA BUSCAR DADOS
// ============================================

function buscarTurmas() {
    try {
        const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('TURMAS_ALUNOS');
        if (!sheet) return [];
        
        const dados = sheet.getDataRange().getValues();
        if (dados.length < 2) return [];
        
        const turmas = new Set();
        for (let i = 1; i < dados.length; i++) {
            const turma = String(dados[i][0]).trim();
            if (turma) turmas.add(turma);
        }
        
        return Array.from(turmas).sort();
    } catch (error) {
        console.error('Erro ao buscar turmas:', error);
        return [];
    }
}

function buscarAlunosPorTurma(turma) {
    try {
        const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('TURMAS_ALUNOS');
        if (!sheet) return [];
        
        const dados = sheet.getDataRange().getValues();
        if (dados.length < 2) return [];
        
        const alunos = [];
        for (let i = 1; i < dados.length; i++) {
            const turmaDado = String(dados[i][0]).trim();
            if (turmaDado === turma) {
                const nome = String(dados[i][1]).trim();
                if (nome) alunos.push(nome);
            }
        }
        
        return alunos.sort();
    } catch (error) {
        console.error('Erro ao buscar alunos:', error);
        return [];
    }
}

function buscarProfessores() {
    try {
        const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('ATRIBUICOES');
        if (!sheet) return [];
        
        const dados = sheet.getDataRange().getValues();
        if (dados.length < 2) return [];
        
        const professores = [];
        const professoresMap = new Map();
        
        for (let i = 1; i < dados.length; i++) {
            const id = String(dados[i][0]).trim();
            const nome = String(dados[i][1]).trim();
            const tipo = String(dados[i][2]).trim() || 'Regente';
            const ativo = String(dados[i][5]).toUpperCase() === 'SIM';
            
            if (id && nome && ativo && !professoresMap.has(id)) {
                professoresMap.set(id, { id, nome, tipo });
            }
        }
        
        return Array.from(professoresMap.values());
    } catch (error) {
        console.error('Erro ao buscar professores:', error);
        return [];
    }
}

function buscarAtribuicoesPorProfessor(professorId) {
    const cabecalho = ['ID_Professor', 'Nome_Professor', 'Tipo', 'Turma', 'Data_Atribuicao', 'Ativo'];
    const atribuicoes = buscarDados('ATRIBUICOES', 'ID_Professor', professorId, cabecalho);
    
    return atribuicoes
        .filter(a => String(a.Ativo).toUpperCase() === 'SIM')
        .map(a => a.Turma);
}

function buscarRelatoriosGerados() {
    try {
        const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('RELATORIOS_GERADOS');
        if (!sheet) return [];
        
        const dados = sheet.getDataRange().getValues();
        if (dados.length < 2) return [];
        
        const cabecalho = dados[0];
        const relatorios = [];
        
        for (let i = 1; i < dados.length; i++) {
            const relatorio = {};
            cabecalho.forEach((col, j) => {
                relatorio[col] = dados[i][j] !== undefined ? dados[i][j] : '';
            });
            
            // Adiciona contagem
            relatorio.Total_Relatorios = dados.length - 1; // Total de linhas menos cabeçalho
            
            relatorios.push(relatorio);
        }
        
        return relatorios;
    } catch (error) {
        console.error('Erro ao buscar relatórios:', error);
        return [];
    }
}

function contarRelatoriosPorStatus(status) {
    try {
        const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('RELATORIOS_GERADOS');
        if (!sheet) return 0;
        
        const dados = sheet.getDataRange().getValues();
        if (dados.length < 2) return 0; // Só tem cabeçalho
        
        const cabecalho = dados[0];
        let colStatus = -1;
        
        cabecalho.forEach((col, index) => {
            if (String(col).trim() === 'Visto') {
                colStatus = index;
            }
        });
        
        if (colStatus === -1) return dados.length - 1; // Retorna total de linhas (menos cabeçalho)
        
        let contador = 0;
        for (let i = 1; i < dados.length; i++) {
            const valorStatus = String(dados[i][colStatus] || '').toUpperCase();
            
            if (!status || status === 'todos') {
                contador++;
            } else if (status === 'vistos' && valorStatus === 'SIM') {
                contador++;
            } else if (status === 'nao_vistos' && valorStatus !== 'SIM') {
                contador++;
            }
        }
        
        return contador;
    } catch (error) {
        console.error('Erro ao contar relatórios:', error);
        return 0;
    }
}

// ============================================
// WEB APP ENDPOINTS
// ============================================

function doGet(e) {
    try {
        let acao = '';
        let payload = {};
        
        if (e && e.parameter) {
            if (e.parameter.data) {
                try {
                    const parsed = JSON.parse(e.parameter.data);
                    acao = parsed.acao;
                    payload = parsed.payload || {};
                } catch (e2) {
                    acao = e.parameter.acao || '';
                    payload = e.parameter;
                }
            } else {
                acao = e.parameter.acao || '';
                payload = e.parameter;
            }
        }
        
        if (!acao) {
            return ContentService
                .createTextOutput(JSON.stringify({ 
                    success: true, 
                    message: 'Sistema de Gestão Escolar - GAS ativo!',
                    timestamp: new Date().toISOString(),
                    versao: '2.0'
                }))
                .setMimeType(ContentService.MimeType.JSON);
        }
        
        let resultado = processarAcao(acao, payload);
        
        return ContentService
            .createTextOutput(JSON.stringify(resultado))
            .setMimeType(ContentService.MimeType.JSON);
            
    } catch (error) {
        console.error('Erro no doGet:', error);
        return ContentService
            .createTextOutput(JSON.stringify({ 
                success: false, 
                error: error.message,
                stack: error.stack
            }))
            .setMimeType(ContentService.MimeType.JSON);
    }
}

function doPost(e) {
    try {
        let acao = '';
        let payload = {};
        
        if (e.postData && e.postData.contents) {
            try {
                const dados = JSON.parse(e.postData.contents);
                acao = dados.acao;
                payload = dados.payload || {};
            } catch (e2) {
                const params = new URLSearchParams(e.postData.contents);
                const dataStr = params.get('data');
                if (dataStr) {
                    const parsed = JSON.parse(dataStr);
                    acao = parsed.acao;
                    payload = parsed.payload || {};
                } else {
                    acao = params.get('acao') || '';
                    payload = Object.fromEntries(params);
                }
            }
        }
        
        if (!acao) {
            return ContentService
                .createTextOutput(JSON.stringify({ 
                    success: false, 
                    error: 'Ação não especificada' 
                }))
                .setMimeType(ContentService.MimeType.JSON);
        }
        
        let resultado = processarAcao(acao, payload);
        
        return ContentService
            .createTextOutput(JSON.stringify(resultado))
            .setMimeType(ContentService.MimeType.JSON);
            
    } catch (error) {
        console.error('Erro no doPost:', error);
        return ContentService
            .createTextOutput(JSON.stringify({ 
                success: false, 
                error: error.message,
                stack: error.stack
            }))
            .setMimeType(ContentService.MimeType.JSON);
    }
}

function processarAcao(acao, payload) {
    console.log(`📥 Processando ação: ${acao}`, payload);
    
    try {
        switch (acao) {
            case 'testar':
                return { 
                    success: true, 
                    message: 'Conexão OK', 
                    timestamp: new Date().toISOString(),
                    versao: '2.0'
                };
                
            case 'atribuirTurma':
                return atribuirTurma(
                    payload.professorId,
                    payload.professorNome,
                    payload.tipo,
                    payload.turma,
                    payload.dataAtribuicao
                );
                
            case 'removerTurma':
                return removerTurma(payload.professorId, payload.turma);
                
            case 'solicitarRelatorio':
                return solicitarRelatorio(
                    payload.turma,
                    payload.aluno,
                    payload.bimestre,
                    payload.professor
                );
                
            case 'enviarRelatorio':
                return enviarRelatorio(
                    payload.turma,
                    payload.aluno,
                    payload.bimestre,
                    payload.professor,
                    payload.notas,
                    payload.media,
                    payload.presenca,
                    payload.observacoes,
                    payload.recomendacao,
                    payload.conteudo
                );
                
            case 'marcarRelatorioComoVisto':
                return marcarRelatorioComoVisto(
                    payload.idRelatorio,
                    payload.turma,
                    payload.aluno,
                    payload.bimestre,
                    payload.professor
                );
                
            case 'salvarNotas':
                return salvarMultiplasNotas(payload.notas || []);
                
            case 'salvarPresenca':
                return salvarPresenca(
                    payload.turma,
                    payload.data,
                    payload.alunosPresentes || [],
                    payload.alunosFaltantes || [],
                    payload.observacao || ''
                );
                
            case 'salvarConteudo':
                return salvarConteudo(
                    payload.turma,
                    payload.data,
                    payload.disciplina,
                    payload.conteudo,
                    payload.objetivos
                );
                
            case 'salvarObservacao':
                return salvarObservacao(
                    payload.turma,
                    payload.data,
                    payload.aluno,
                    payload.observacao,
                    payload.tipo
                );
                
            case 'buscarTurmas':
                return { success: true, turmas: buscarTurmas() };
                
            case 'buscarAlunos':
                return { success: true, alunos: buscarAlunosPorTurma(payload.turma) };
                
            case 'buscarProfessores':
                return { success: true, professores: buscarProfessores() };
                
            case 'buscarAtribuicoes':
                return { success: true, atribuicoes: buscarAtribuicoesPorProfessor(payload.professorId) };
                
            case 'buscarRelatoriosGerados':
                return { success: true, relatorios: buscarRelatoriosGerados() };
                
            case 'contarRelatorios':
                return { 
                    success: true, 
                    total: contarRelatoriosPorStatus('todos'),
                    vistos: contarRelatoriosPorStatus('vistos'),
                    nao_vistos: contarRelatoriosPorStatus('nao_vistos')
                };
                
            default:
                return { success: false, error: 'Ação não reconhecida: ' + acao };
        }
    } catch (error) {
        console.error(`Erro ao processar ação ${acao}:`, error);
        return { 
            success: false, 
            error: error.message,
            stack: error.stack
        };
    }
}

// ============================================
// FUNÇÃO DE TESTE
// ============================================

function testarSistema() {
    try {
        const resultado = {
            timestamp: new Date().toISOString(),
            spreadsheets: {},
            abas: {}
        };
        
        // Verifica se a planilha existe
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        resultado.spreadsheets.nome = ss.getName();
        resultado.spreadsheets.id = SPREADSHEET_ID;
        
        // Lista todas as abas
        const sheets = ss.getSheets();
        resultado.abas.total = sheets.length;
        resultado.abas.nomes = sheets.map(s => s.getName());
        
        // Verifica abas essenciais
        const abasEssenciais = [
            'TURMAS_ALUNOS',
            'ATRIBUICOES',
            'CONTEUDOS',
            'OBSERVACOES',
            'RELATORIOS_SOLICITADOS',
            'RELATORIOS_GERADOS'
        ];
        
        resultado.abas.essenciais = {};
        abasEssenciais.forEach(aba => {
            resultado.abas.essenciais[aba] = sheets.some(s => s.getName() === aba);
        });
        
        // Verifica abas de presença
        const abasPresenca = sheets.filter(s => s.getName().startsWith('PRESENCA_'));
        resultado.abas.presenca = abasPresenca.map(s => s.getName());
        
        // Verifica abas de notas
        const abasNotas = sheets.filter(s => s.getName().startsWith('NOTAS_BIMESTRAIS_'));
        resultado.abas.notas = abasNotas.map(s => s.getName());
        
        return resultado;
    } catch (error) {
        console.error('Erro ao testar sistema:', error);
        return {
            success: false,
            error: error.message,
            timestamp: new Date().toISOString()
        };
    }
}

// ============================================
// FUNÇÃO PARA INICIALIZAR ESTRUTURA DA PLANILHA
// ============================================

function inicializarEstrutura() {
    try {
        const resultados = [];
        
        // 1. Cria aba TURMAS_ALUNOS
        const cabecalhoTurmas = ['Turma', 'Nome_Aluno', 'Ativo'];
        const sheetTurmas = getOrCreateSheet('TURMAS_ALUNOS', cabecalhoTurmas);
        resultados.push({ aba: 'TURMAS_ALUNOS', criada: true });
        
        // 2. Cria aba ATRIBUICOES
        const cabecalhoAtribuicoes = ['ID_Professor', 'Nome_Professor', 'Tipo', 'Turma', 'Data_Atribuicao', 'Ativo'];
        const sheetAtribuicoes = getOrCreateSheet('ATRIBUICOES', cabecalhoAtribuicoes);
        resultados.push({ aba: 'ATRIBUICOES', criada: true });
        
        // 3. Cria aba CONTEUDOS
        const cabecalhoConteudos = ['Turma', 'Data', 'Disciplina', 'Conteudo', 'Objetivos'];
        const sheetConteudos = getOrCreateSheet('CONTEUDOS', cabecalhoConteudos);
        resultados.push({ aba: 'CONTEUDOS', criada: true });
        
        // 4. Cria aba OBSERVACOES
        const cabecalhoObservacoes = ['Turma', 'Data', 'Aluno', 'Observacao', 'Tipo'];
        const sheetObservacoes = getOrCreateSheet('OBSERVACOES', cabecalhoObservacoes);
        resultados.push({ aba: 'OBSERVACOES', criada: true });
        
        // 5. Cria aba RELATORIOS_SOLICITADOS
        const cabecalhoSolicitacoes = [
            'ID_Solicitacao', 
            'Data_Solicitacao', 
            'Turma', 
            'Aluno', 
            'Bimestre', 
            'Professor', 
            'Status', 
            'Data_Conclusao', 
            'Observacoes'
        ];
        const sheetSolicitacoes = getOrCreateSheet('RELATORIOS_SOLICITADOS', cabecalhoSolicitacoes);
        resultados.push({ aba: 'RELATORIOS_SOLICITADOS', criada: true });
        
        // 6. Cria aba RELATORIOS_GERADOS
        const cabecalhoRelatorios = [
            'ID_Relatorio',
            'Data_Geracao',
            'Turma',
            'Aluno',
            'Bimestre',
            'Professor',
            'Notas',
            'Media',
            'Presenca',
            'Observacoes',
            'Recomendacao',
            'Conteudo',
            'Visto',
            'Data_Visto'
        ];
        const sheetRelatorios = getOrCreateSheet('RELATORIOS_GERADOS', cabecalhoRelatorios);
        resultados.push({ aba: 'RELATORIOS_GERADOS', criada: true });
        
        return {
            success: true,
            message: 'Estrutura inicializada com sucesso!',
            resultados: resultados
        };
    } catch (error) {
        console.error('Erro ao inicializar estrutura:', error);
        return {
            success: false,
            error: error.message
        };
    }
}

// ============================================
// FUNÇÃO PARA CONFIGURAÇÃO INICIAL
// ============================================

function setup() {
    try {
        // Inicializa estrutura
        const resultado = inicializarEstrutura();
        
        // Testa o sistema
        const teste = testarSistema();
        
        return {
            success: true,
            message: 'Sistema configurado com sucesso!',
            estrutura: resultado,
            teste: teste
        };
    } catch (error) {
        console.error('Erro no setup:', error);
        return {
            success: false,
            error: error.message
        };
    }
}

