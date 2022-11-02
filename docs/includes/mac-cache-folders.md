Os suplementos geralmente são armazenados em cache no Office no Mac por motivos de desempenho. Normalmente, o cache será limpo quando o suplemento for recarregado. Se houver mais de um suplemento no mesmo documento, é provável que o processo de limpeza automática do cache ao recarregar não seja confiável.

### <a name="use-the-personality-menu-to-clear-the-cache"></a>Usar o menu de personalidade para limpar o cache

Você pode limpar o cache usando o menu personalidade de qualquer suplemento do painel de tarefas. No entanto, como o menu de personalidade não tem suporte nos suplementos do Outlook, você pode tentar a opção de [limpar o cache manualmente](#clear-the-cache-manually) se estiver usando o Outlook.

- Escolha o menu personalidade. Em seguida, escolha **Limpar Cache da Web**.
    > [!NOTE]
    > Você deve executar o macOS Versão 10.13.6 ou posterior para ver o menu de personalidade.

    ![Captura de tela da opção limpar cache da web em um menu de personalidade.](../images/mac-clear-cache-menu.png)

### <a name="clear-the-cache-manually"></a>Limpar o cache manualmente

Você também pode limpar o cache manualmente ao excluir o conteúdo na pasta `~/Library/Containers/com.Microsoft.OsfWebHost/Data/`. Procure essa pasta por meio do terminal.

> [!NOTE]
> Se essa pasta não existir, verifique se há as pastas a seguir por meio do terminal e, se encontradas, exclua o conteúdo da pasta.
>
> - `~/Library/Containers/com.microsoft.{host}/Data/Library/Caches/` onde `{host}` é o aplicativo do Office (por exemplo, `Excel`)
> - `~/Library/Containers/com.microsoft.{host}/Data/Library/Application Support/Microsoft/Office/16.0/Wef/` onde `{host}` é o aplicativo do Office (por exemplo, `Excel`)
> - `~/Library/Containers/com.microsoft.Office365ServiceV2/Data/Caches/com.microsoft.Office365ServiceV2/`
> - `~/Library/Containers/com.microsoft.Office365ServiceV2/Data/Library/Caches/com.microsoft.Office365ServiceV2/`
>
> Para procurar essas pastas por meio do Finder, você deve definir o Finder para mostrar arquivos ocultos. O Localizador exibe as pastas dentro do diretório **Contêineres** pelo nome do produto, como **o Microsoft Excel** em vez de **com.microsoft.Excel**.