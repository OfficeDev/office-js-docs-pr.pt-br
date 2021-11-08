Use o procedimento a seguir para instalar uma versão de assinatura Office que usa o webview Versão Prévia do Microsoft Edge (EdgeHTML) para executar os complementos ou uma versão que usa o Internet Explorer (Trident).

1. Em qualquer Office aplicativo, abra a guia **Arquivo** na faixa de opções e selecione Office **Conta** ou **Conta.** Selecione o **botão Sobre nome do _host_** (por exemplo, **Sobre o Word**).
1. Na caixa de diálogo que é aberta, encontre o número de com build completo xx.x.xxxxx.xxxxx e faça uma cópia dele em algum lugar.
1. Baixe e instale a [ferramenta Office implantação.](https://www.microsoft.com/download/details.aspx?id=49117)
1. Na pasta onde você instalou a ferramenta (onde o arquivo está localizado), crie um arquivo de texto com o nome e `setup.exe` `config.xml` adicione o seguinte conteúdo.

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

    - Para instalar uma versão que usa o Edge Legacy, altere-a para `16.0.11929.20946` .
    - Para instalar uma versão que usa o Internet Explorer, altere-a para `16.0.10730.20348` .

1. Opcionalmente, altere o valor de para instalar Office de 32 bits e altere o valor conforme necessário para instalar o Office `OfficeClientEdition` `"32"` em um idioma `Language ID` diferente.
1. Abra um prompt de comando *como administrador*.
1. Navegue até a pasta com `setup.exe` os arquivos `config.xml` e.
1. Execute o seguinte comando.

    ```command&nbsp;line
    setup.exe /configure config.xml
    ```

    Este comando instala o Office. O processo pode levar vários minutos.

1. [Limpar o Office cache](../testing/clear-cache.md).

> [!IMPORTANT]
> Após a instalação, certifique-se de desativar a atualização automática do Office, para que o Office não seja atualizado para uma versão que não use webview com a qual você deseja trabalhar antes de concluir o uso. **Isso pode acontecer em minutos após a instalação.** Siga estas etapas.
>
> 1. Inicie qualquer Office aplicativo e abra um novo documento.
> 1. Abra a **guia Arquivo** na faixa de opções e selecione Office **Conta** ou **Conta.**
> 1. Na coluna **Informações do Produto,** selecione **Opções de** Atualização e, em seguida, **selecione Desabilitar Atualizações**. Se essa opção não estiver disponível, a Office já está configurada para não ser atualizada automaticamente.

Quando terminar de usar a versão antiga do Office, reinstale sua versão mais recente editando o arquivo e alterando o para o número de com build que você `config.xml` `Version` copiou anteriormente. Em seguida, repita `setup.exe /configure config.xml` o comando em um prompt de comando do administrador. Opcionalmente, reabilitar atualizações automáticas.
