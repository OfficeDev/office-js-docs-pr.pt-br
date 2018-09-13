Neste tutorial, comece configurando seu projeto de desenvolvimento. 

> [!NOTE]
> Esta página descreve uma etapa individual de um tutorial de suplemento do Word. Se você chegou aqui por meio dos resultados de mecanismos de pesquisa ou por outro link direto, acesse a página de introdução do [tutorial de suplemento do Word](../tutorials/word-tutorial.yml) para começar pelo início.

> [!TIP]
> Leia [Compilar seu primeiro suplemento do Word](../quickstarts/word-quickstart.md?tabs=visual-studio-code), se ainda não tiver lido. Em particular, você deve saber fazer o sideload de um suplemento do Word para teste.

## <a name="prerequisites"></a>Pré-requisitos

Para usar este tutorial, você precisa instalar o seguinte. 

- Word 2016, versão 1711 (build 8730.1000 do Clique para Executar) ou posterior. Talvez você precise ser um participante do programa Office Insider para ter essa versão. Para saber mais, confira a página [Seja um Office Insider](https://products.office.com/office-insider?tab=tab-1).
- [Nó e npm](https://nodejs.org/en/) 
- [Git Bash](https://git-scm.com/downloads) (ou outro cliente Git)

## <a name="setup"></a>Configurar

1. Clone o repositório do GitHub com o [Tutorial de suplemento do Word](https://github.com/OfficeDev/Word-Add-in-Tutorial).
2. Abra uma janela bash do Git ou um prompt de sistema habilitado para Node.JS e navegue para a pasta **Iniciar** do projeto.
3. Execute o comando `npm install` para instalar as ferramentas e bibliotecas listadas no arquivo package.json. 
4. Execute as etapas em [Adicionar certificados autoassinados como certificado raiz confiáveis](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) para confiar no certificado do sistema operacional do seu computador de desenvolvimento.

