Neste tutorial, comece configurando seu projeto de desenvolvimento. 

> [!NOTE]
> Esta página descreve uma etapa individual do tutorial de suplemento do Excel. Se você chegou aqui por meio dos resultados de mecanismos de pesquisa ou via outro link direto, acesse a página de Introdução do [tutorial de suplemento do Excel](../tutorials/excel-tutorial.yml) para começá-lo do início.

## <a name="prerequisites"></a>Pré-requisitos

Para usar este tutorial, você precisa instalar o seguinte. 

- Excel 2016, versão 1711 (build 8730.1000 do Clique para Executar) ou posterior. Talvez você precise ser um participante do programa Office Insider para ter essa versão. Para saber mais, confira a página [Seja um Office Insider](https://products.office.com/office-insider?tab=tab-1).
- [Nó e npm](https://nodejs.org/en/) 
- [Git Bash](https://git-scm.com/downloads) (ou outro cliente Git)

## <a name="setup"></a>Configurar

1. Clone o repositório do GitHub com o [Tutorial de suplemento do Excel](https://github.com/OfficeDev/Excel-Add-in-Tutorial).
2. Abra uma janela Git bash ou um prompt de sistema habilitado para Node.JS e navegue para a pasta **Iniciar** do projeto.
3. Execute o comando `npm install` para instalar as ferramentas e bibliotecas listadas no arquivo package.json. 
4. Execute as etapas em [Adicionar certificados autoassinados como certificado raiz confiáveis](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) para confiar no certificado do sistema operacional do seu computador de desenvolvimento.

