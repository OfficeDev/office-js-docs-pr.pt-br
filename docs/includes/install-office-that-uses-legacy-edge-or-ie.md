Use o procedimento a seguir para instalar uma versão do Office (baixada de uma assinatura do Microsoft 365) que usa o Versão Prévia do Microsoft Edge Webview (EdgeHTML) para executar suplementos ou uma versão que usa o Internet Explorer (Trident).

1. Em qualquer aplicativo do Office, abra a guia **Arquivo** na faixa de opções e selecione **Conta ou Conta do Office**. Selecione o botão **Sobre _nome do host_** (por exemplo, **Sobre o Word**).
1. Na caixa de diálogo aberta, localize o número completo de build xx.x.xxxxx.xxxxx e faça uma cópia dele em algum lugar.
1. Baixar a [Ferramenta de Implantação do Office](https://www.microsoft.com/download/details.aspx?id=49117).
1. Execute o arquivo baixado para extrair a ferramenta. Você deve escolher onde instalar a ferramenta.
1. Na pasta em que você instalou a ferramenta (onde o `setup.exe` arquivo está localizado), crie um arquivo de texto com o nome `config.xml` e adicione o conteúdo a seguir.

    ```xml
    <Configuration>
      <Add OfficeClientEdition="64" Channel="SemiAnnual" Version="16.0.xxxxx.xxxxx">
        <Product ID="O365ProPlusRetail">
          <Language ID="en-us" />
        </Product>
      </Add>
    </Configuration>
    ```

1. Altere o `Version` valor.

    - Para instalar uma versão que usa o Edge Legacy, altere-a para `16.0.11929.20946`.
    - Para instalar uma versão que usa o Internet Explorer, altere-a para `16.0.10730.20348`.

1. Opcionalmente, altere o valor de `OfficeClientEdition` para `"32"` instalar o Office de 32 bits e altere o valor conforme necessário para instalar o `Language ID` Office em um idioma diferente.
1. Abra um prompt de comando *como administrador*.
1. Navegue até a pasta com os `setup.exe` arquivos e `config.xml` .
1. Execute o seguinte comando:

    ```command&nbsp;line
    setup.exe /configure config.xml
    ```

    Esse comando instala o Office. Esse processo pode demorar alguns minutos.

1. [Desmarque o cache do Office](../testing/clear-cache.md).

> [!IMPORTANT]
> Após a instalação, desative a atualização automática do Office para que o Office não seja atualizado para uma versão que não use a visão da Web com a qual você deseja trabalhar antes de concluir o uso. **Isso pode acontecer em poucos minutos de instalação.** Siga estas etapas.
>
> 1. Inicie qualquer aplicativo do Office e abra um novo documento.
> 1. Abra a guia **Arquivo** na faixa de opções e selecione **Conta do Office** ou **Conta**.
> 1. Na coluna **Informações do Produto**, selecione **Opções de Atualização** e selecione **Desabilitar Atualizações**. Se essa opção não estiver disponível, o Office já está configurado para não atualizar automaticamente.

Quando terminar de usar a versão antiga do Office, reinstale sua versão mais recente editando o `config.xml` arquivo e alterando o para o `Version` número de build copiado anteriormente. Em seguida, repita o `setup.exe /configure config.xml` comando em um prompt de comando de administrador. Opcionalmente, habilite novamente as atualizações automáticas.
