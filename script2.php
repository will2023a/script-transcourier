<?php

require 'vendor/autoload.php';

use Shuchkin\SimpleXLSXGen;
use Shuchkin\SimpleXLSX;

function extrair_dados($arquivo)
{
    echo "Lendo o arquivo: $arquivo\n";
    $conteudo = file_get_contents($arquivo);
    if (!$conteudo) {
        echo "Erro ao ler o arquivo: $arquivo\n";
        return [];
    }

    $dados = [];

    // Pega o nome da transportadora
    preg_match('/<strong>([^<]+)<\/strong>/', $conteudo, $transportadora);
    $dados['transportadora'] = $transportadora[1] ?? 'N/A';

    // Pega o endereço da transportadora
    preg_match('/(RUA [^<]+<\/td>)/', $conteudo, $endereco);
    $dados['endereco'] = isset($endereco[1]) ? strip_tags($endereco[1]) : 'N/A';

    // Pega a cidade da transportadora
    preg_match('/([A-Z\s\-]+)\s*\-\s*([A-Z]{2})/', $conteudo, $cidade);
    $dados['cidade'] = isset($cidade[1]) ? trim($cidade[1]) . ' - ' . trim($cidade[2]) : 'N/A';

    // Pega os dados numéricos: AWB, Volume, Peso, Peso Cubado, Peso Taxado, Valor
    preg_match_all('/<td[^>]*>([\d.,]+)<\/td>/', $conteudo, $matches);

    if (empty($matches[1])) {
        echo "Nenhum dado de volume, peso ou destino encontrado no arquivo: $arquivo\n";
        return $dados;
    }

    $dados['awb'] = $matches[1][0] ?? 'N/A';
    $dados['volume'] = $matches[1][1] ?? 'N/A';
    $dados['peso'] = $matches[1][2] ?? 'N/A';
    $dados['peso_cubado'] = $matches[1][3] ?? 'N/A';
    $dados['peso_taxado'] = $matches[1][4] ?? 'N/A';
    $dados['valor'] = $matches[1][5] ?? 'N/A';

    // Captura o destino (não numérico, por isso separado)
    preg_match('/<td[^>]*>([A-Z]+)<\/td>\s*<td[^>]*>[\d.,]+<\/td>/', $conteudo, $destino);
    $dados['destino'] = isset($destino[1]) ? trim($destino[1]) : 'N/A';

    return $dados;
}

// Perguntar ao usuário o caminho da pasta
$caminho_pasta = readline("Por favor, insira o caminho da pasta onde estão os arquivos PHP: ");

// Verificar se o diretório inserido pelo usuário é válido
if (!is_dir($caminho_pasta)) {
    echo "O caminho fornecido não é um diretório válido.\n";
    exit;
}

// Garantir que o caminho termine com uma barra
if (substr($caminho_pasta, -1) !== DIRECTORY_SEPARATOR) {
    $caminho_pasta .= DIRECTORY_SEPARATOR;
}

echo "Procurando arquivos PHP no diretório: $caminho_pasta\n";
$arquivos = glob($caminho_pasta . '*.php');

if (empty($arquivos)) {
    echo "Nenhum arquivo PHP encontrado no diretório: $caminho_pasta\n";
    exit;
}

// Iniciar com o cabeçalho
$linhas = [
    ['Transportadora', 'Endereço', 'Cidade', 'AWB', 'Volume', 'Peso', 'Peso Cubado', 'Peso Taxado', 'Valor', 'Destino']
];

echo "Iniciando a extração de dados dos arquivos...\n";

foreach ($arquivos as $arquivo) {
    $dados = extrair_dados($arquivo);
    if (!empty($dados)) {
        $linhas[] = [
            $dados['transportadora'],
            $dados['endereco'],
            $dados['cidade'],
            $dados['awb'],
            $dados['volume'],
            $dados['peso'],
            $dados['peso_cubado'],
            $dados['peso_taxado'],
            $dados['valor'],
            $dados['destino']
        ];
    } else {
        echo "Nenhum dado válido extraído do arquivo: $arquivo\n";
    }
}

// Usar o nome da pasta para gerar o nome do arquivo
$nome_pasta = basename(rtrim($caminho_pasta, DIRECTORY_SEPARATOR));
$data_hora = date('Ymd_His');
$arquivo_excel = "{$nome_pasta}_$data_hora.xlsx";

echo "Gerando o arquivo Excel...\n";

// Gerar o arquivo Excel
$xlsx = SimpleXLSXGen::fromArray($linhas);
$xlsx->saveAs($arquivo_excel);

echo "Arquivo Excel gerado com sucesso: $arquivo_excel\n";
