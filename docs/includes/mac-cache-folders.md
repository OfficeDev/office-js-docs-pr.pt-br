Os suplementos são frequentemente armazenados em cache no Office para Mac, por motivos de desempenho. Normalmente, o cache será limpo quando o suplemento for recarregado. Se houver mais de um suplemento no mesmo documento, o processo de limpeza automática do cache no recarregamento poderá não ser confiável.

Você pode limpar o cache usando o menu de personalidade de qualquer suplemento do painel de tarefas.
- Escolha o menu de personalidade. Em seguida, escolha **Limpar cache da Web**.
    > [!NOTE]
    > Você deve executar o macOS versão 10.13.6 ou posterior para ver o menu de personalidade.
    
    ![Captura de tela da opção Limpar cache da Web no menu personalidade.](../images/mac-clear-cache-menu.png)

Você também pode limpar o cache manualmente excluindo o conteúdo da `~/Library/Containers/com.Microsoft.OsfWebHost/Data/` pasta.

> [!NOTE]
> Se essa pasta não existir, verifique as seguintes pastas e, se encontrar, exclua o conteúdo da pasta:
>    - `~/Library/Containers/com.microsoft.{host}/Data/Library/Caches/`onde `{host}` é o host do Office (por exemplo `Excel`,)
>    - `com.microsoft.Office365ServiceV2/Data/Caches/com.microsoft.Office365ServiceV2/`
