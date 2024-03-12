using System.Runtime.Intrinsics.X86;
using ConsoleApp1;

using var escritor = new EscritorExcel(@"C:\Users\andrew.maia\Desktop\QAR\QarSaude_01_Dinamico.xlsx");


//Aba Estrategia
const string 
    consultor = "B7", 
    razaoSocialEstipulante="B8",
    estipulanteCnpj="J8",
    operadoraAtual="B10",
    adaptadoLei="D10",
    tempoContrato="G10",
    agregados="I10",
    agregadosComplemento="J10",
    aniversarioContrato="B11",
    breakEven="D11",
    afastados="I11",
    afastadosComplemento="J11",
    possuiMulta="B12",
    regraMulta="D12",
    aposentadosPor="I12",    
    aposentadosPorComplemento="J12",    
    relatorioSinistralidade="B13",
    percentualSinistralidade="D13",
    gestantes="I13",
    gestantesComplemento="J13",    
    modalidadeContrato="B14",
    haDependentes="D14",
    cronicos="I14",    
    cronicosComplemento="J14",
    haReeembolso="B15",    
    remidos="I15",    
    remidosComplemento="J15",    
    coparticipacao="B16",
    regraCoparticacao="D16",
    inativos="I16",
    inativosComplemento="J16",
    regraUpDownGrade="B17",
    prestadorServico="I17",
    prestadorServicoComplemento="J17",
    contribuicaoTitular="B18",    
    contribuicaoTitularEmpresa="E18",        
    estagiarios="I18",
    estagiariosComplemento="J18",
    contribuicaoDependente="B19",    
    contribuicaoDependenteEmpresa="E19",        
    homeCare="I19",
    homeCareComplemento="J19",
    elegibilidade="B20",    
    elegibilidadeCnpj="J20";

escritor.AlterarValorCelula(consultor,"José da Silva");
escritor.AlterarValorCelula(razaoSocialEstipulante,"Razao social Ficticia");
escritor.AlterarValorCelula(estipulanteCnpj , "91.786.878/0001-50");
escritor.AlterarValorCelula(operadoraAtual , "Bradesco");
escritor.AlterarValorCelula(adaptadoLei , "Sim");
escritor.AlterarValorCelula(tempoContrato , "2 anos");
escritor.AlterarValorCelula(agregados , "Sim");
escritor.AlterarValorCelula(agregadosComplemento , "Complemento");
escritor.AlterarValorCelula(aniversarioContrato , new DateTime(2025,01,31));
escritor.AlterarValorCelula(breakEven , "Break Even");
escritor.AlterarValorCelula(afastados , "Não");
escritor.AlterarValorCelula(afastadosComplemento , "Complemento");
escritor.AlterarValorCelula(possuiMulta , "Sim");
escritor.AlterarValorCelula(regraMulta , "Alguma Regra");
escritor.AlterarValorCelula(aposentadosPor , "Sim");
escritor.AlterarValorCelula(aposentadosPorComplemento , "Complemento");
escritor.AlterarValorCelula(relatorioSinistralidade , "Sim");
escritor.AlterarValorCelula(percentualSinistralidade , 0.2d);
escritor.AlterarValorCelula(gestantes , "Sim");
escritor.AlterarValorCelula(gestantesComplemento , "Complemento");
escritor.AlterarValorCelula(modalidadeContrato , "Compulsório");
escritor.AlterarValorCelula(haDependentes , "Sim");
escritor.AlterarValorCelula(cronicos , "Não");
escritor.AlterarValorCelula(cronicosComplemento , "Complemento");
escritor.AlterarValorCelula(haReeembolso , "Sim");
escritor.AlterarValorCelula(remidos , "Não");
escritor.AlterarValorCelula(remidosComplemento , "Complemento");
escritor.AlterarValorCelula(coparticipacao , "Sim");
escritor.AlterarValorCelula(regraCoparticacao , "Alguma Regra");
escritor.AlterarValorCelula(inativos , "Sim");
escritor.AlterarValorCelula( inativosComplemento, "Complementos");
escritor.AlterarValorCelula( regraUpDownGrade, "Alguma regra");
escritor.AlterarValorCelula( prestadorServico, "Não");
escritor.AlterarValorCelula( prestadorServicoComplemento, "Complemento");
escritor.AlterarValorCelula(contribuicaoTitular , 0.1);
escritor.AlterarValorCelula(contribuicaoTitularEmpresa , 0.9);
escritor.AlterarValorCelula( estagiarios, "Sim");
escritor.AlterarValorCelula( estagiariosComplemento, "Complemento");
escritor.AlterarValorCelula(contribuicaoDependente , 0.3);
escritor.AlterarValorCelula(contribuicaoDependenteEmpresa , 0.7);
escritor.AlterarValorCelula( homeCare, "Sim");
escritor.AlterarValorCelula( homeCareComplemento, "Complemento");
escritor.AlterarValorCelula( elegibilidade, "Sim");
escritor.AlterarValorCelula( elegibilidadeCnpj, "91.786.878/0001-50");


//Estratégia

//Sub Estipulante
const string linhaReferenciaSubEstipulante  ="A34";
var subEstipulantes = Mock.MockarSubEstimulantes();

Linha linhaReferencia=escritor.ObterLinha(linhaReferenciaSubEstipulante);
uint indexLinha = linhaReferencia.Index;

foreach(var subEstipulante in subEstipulantes){
    indexLinha++;
    var linha  =escritor.ClonarLinha(linhaReferencia,indexLinha);
    linha.AlterarCelula(subEstipulante.RazaoSocial,1);
    linha.AlterarCelula(subEstipulante.CNPJ,3);    
}

//Aba BASE DE DADOS  ESTUDOS
escritor.AlterarWorksheet("BASE DE DADOS  ESTUDOS");
const string linhaReferenciaBaseDados  ="A2";
var pessoas = Mock.MockarBaseDadosSaude();

linhaReferencia=escritor.ObterLinha(linhaReferenciaBaseDados);
indexLinha = linhaReferencia.Index;

foreach(var pessoa in pessoas){
    var linha  = escritor.ClonarLinha(linhaReferencia,indexLinha++);

    int indexCelula=0;
    linha.AlterarCelula(pessoa.Empresa,indexCelula++);
    linha.AlterarCelula(pessoa.CNPJ,indexCelula++);
    linha.AlterarCelula(pessoa.Sexo,indexCelula++);
    linha.AlterarCelula(pessoa.Identificacao,indexCelula++);
    linha.AlterarCelula(pessoa.DataNascimento,indexCelula++);
    linha.AlterarCelula(pessoa.Idade,indexCelula++);
    linha.AlterarCelula(pessoa.FaixaEtaria,indexCelula++);
    linha.AlterarCelula(pessoa.Parentesto,indexCelula++);
    linha.AlterarCelula(pessoa.Situacao,indexCelula++);
    linha.AlterarCelula(pessoa.CID,indexCelula++);
    linha.AlterarCelula(pessoa.Municipio,indexCelula++);
    linha.AlterarCelula(pessoa.UF,indexCelula++);
    linha.AlterarCelula(pessoa.Operadora,indexCelula++);
    linha.AlterarCelula(pessoa.Plano,indexCelula++);
    linha.AlterarCelula(pessoa.ValorAtual,indexCelula++);
}

 escritor.GerarArquivo(@"C:\Users\andrew.maia\Desktop\QAR\QarSaude_03_Preenchido.xlsx");




