Use o procedimento a seguir para instalar uma versão do Office de assinatura do Microsoft 365 que usa o modo de exibição da Web do Versão Prévia do Microsoft Edge (EdgeHTML) para executar suplementos ou uma versão que usa o Internet Explorer (Trident).

1. Em qualquer aplicativo do Office, abra **a guia Arquivo** na faixa de opções e selecione Conta **do Office** ou **Conta**. Selecione o **botão _Sobre nome do host_** (por exemplo, **Sobre o Word**).
1. Na caixa de diálogo que é aberta, localize o número de build xx.x.xxxxx.xxxxx completo e faça uma cópia dele em algum lugar.
1. Baixar a [Ferramenta de Implantação do Office](https://www.microsoft.com/download/details.aspx?id=49117).
1. Execute o arquivo baixado para extrair a ferramenta. Você será solicitado a escolher onde instalar a ferramenta.
1. Na pasta em que você instalou a `config.xml` ferramenta (`setup.exe`onde o arquivo está localizado), crie um arquivo de texto com o nome e adicione o conteúdo a seguir.

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

1. Opcionalmente, altere `OfficeClientEdition` `"32"` o valor para instalar o Office `Language ID` de 32 bits e altere o valor conforme necessário para instalar o Office em um idioma diferente.
1. Abra um prompt de *comando como administrador*.
1. Navegue até a pasta com os `setup.exe` arquivos `config.xml` e os arquivos.
1. Execute o seguinte comando:

    ```command&nbsp;line
    setup.exe /configure config.xml
    ```

    Este comando instala o Office. Esse processo pode demorar alguns minutos.

1. [Limpe o cache do Office](../testing/clear-cache.md).

> [!IMPORTANT]
> Após a instalação, certifique-se de desativar a atualização automática do Office para que o Office não seja atualizado para uma versão que não use o modo de exibição da Web com o qual você deseja trabalhar antes de concluir o uso. **Isso pode acontecer em minutos após a instalação.** Siga estas etapas.
>
> 1. Inicie qualquer aplicativo do Office e abra um novo documento.
> 1. Abra a **guia Arquivo** na faixa de opções e selecione Conta **ou** **Conta** do Office.
> 1. Na coluna **Informações do Produto**, selecione **Opções de** Atualização e, em seguida, **selecione Desabilitar Atualizações**. Se essa opção não estiver disponível, o Office já estará configurado para não ser atualizado automaticamente.

Quando terminar de usar a versão antiga do Office, reinstale `config.xml` `Version` sua versão mais recente editando o arquivo e alterando o número de build que você copiou anteriormente. Em seguida, repita `setup.exe /configure config.xml` o comando em um prompt de comando do administrador. Opcionalmente, reabilitar atualizações automáticas.
