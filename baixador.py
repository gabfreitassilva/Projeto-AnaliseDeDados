import asyncio
import pathlib
from getpass import getpass
from playwright.async_api import Playwright, async_playwright

user = input("Digite o seu usuário do sistema: ")
password = getpass("Digite a sua senha: ") # coleta a senha sem mostrar os caracteres

async def run(playwright: Playwright) -> None:
    browser = await playwright.firefox.launch(headless=True)
    context = await browser.new_context()
    page = await context.new_page()
    
    # Login no sistema
    print("Realizando login! 🔒")
    await page.goto("http://simg.metrofor.ce.gov.br/")
    await page.get_by_role("textbox", name="Usuario:").click()
    await page.get_by_role("textbox", name="Usuario:").fill(user)
    await page.get_by_role("textbox", name="Senha:").click()
    await page.get_by_role("textbox", name="Senha:").fill(password)
    await page.get_by_role("textbox", name="Senha:").press("Enter")
    print("Login realizado! 🔓")

    # Baixar SAF's
    print("\nRealizando download de SAF's...")
    await page.get_by_role("link", name=" Relatórios ").click()
    await page.get_by_role("link", name="Relatório de Formulários").click()
    await page.get_by_role("link", name="SAF").click()
    await page.get_by_role("tab", name="Data").locator("button").click()
    await page.locator("input[name=\"dataPartirAbertura\"]").click()
    await page.locator("input[name=\"dataPartirAbertura\"]").fill("01102025")
    await page.locator("input[name=\"dataAteAbertura\"]").click()
    await page.locator("input[name=\"dataAteAbertura\"]").fill("31102025")

    # Espera começar o download
    async with page.expect_download() as download_info:
        # Faz a ação que inicia o download
        await page.get_by_role("button", name=" Excel Completo").click()
    download = await download_info.value
    # Obtém o caminho da pasta de Downloads do usuário atual
    downloads_path = pathlib.Path.home() / "Downloads" / "Documentos - Farol"
    await download.save_as(downloads_path / download.suggested_filename)    
    print("Download de SAF's realizado!")

    # Baixar OSM's
    print("\nRelizando download de OSM's...")
    await page.get_by_role("link", name=" Relatórios ").click()
    await page.get_by_role("link", name="Relatório de Formulários").click()
    await page.get_by_role("link", name="OSM").click()
    await page.get_by_role("tab", name="Datas e Status").locator("button").click()
    await page.locator("input[name=\"dataPartirOsmAbertura\"]").click()
    await page.locator("input[name=\"dataPartirOsmAbertura\"]").fill("01102025")
    await page.locator("input[name=\"dataAteOsmAbertura\"]").click()
    await page.locator("input[name=\"dataAteOsmAbertura\"]").fill("31102025")

    # Espera começar o download
    async with page.expect_download() as download_info:
        # Faz a ação que inicia o download
        await page.get_by_role("button", name=" Excel Completo").click()
    download = await download_info.value
    await download.save_as(downloads_path / download.suggested_filename)    
    print("Download de OSM's realizado!")

    # Baixar SSP's
    print("\nRelizando download de SSP's...")
    await page.get_by_role("link", name=" Relatórios ").click()
    await page.get_by_role("link", name="Relatório de Formulários").click()
    await page.get_by_role("link", name="SSP").click()
    await page.get_by_role("tab", name="Status").locator("button").click()
    await page.locator("input[name=\"dataPartirAbertura\"]").click()
    await page.locator("input[name=\"dataPartirAbertura\"]").fill("01102025")    # Modifique a data conforme necessário
    await page.locator("input[name=\"dataAteAbertura\"]").click()
    await page.locator("input[name=\"dataAteAbertura\"]").fill("31102025")       # Modifique a data conforme necessário
    await page.get_by_role("tab", name="Serviço").locator("button").click()
    await page.locator("select[name=\"grupoSistemaPesquisa\"]").select_option("22")

    # Espera começar o download
    async with page.expect_download() as download_info:
        # Faz a ação que inicia o download
        await page.get_by_role("button", name=" Excel Completo").click()
    download = await download_info.value
    await download.save_as(downloads_path / download.suggested_filename)
    print("Download de SSP's realizado!")

    # Baixar OSP's
    print("\nRelizando download de OSP's...")
    await page.get_by_role("link", name=" Relatórios ").click()
    await page.get_by_role("link", name="Relatório de Formulários").click()
    await page.get_by_role("link", name="OSP").click()
    await page.get_by_role("tab", name="Datas e Status").locator("button").click()
    await page.locator("input[name=\"dataPartirOspAbertura\"]").click()
    await page.locator("input[name=\"dataPartirOspAbertura\"]").fill("01102025") # Modifique a data conforme necessário
    await page.locator("input[name=\"dataAteOspAbertura\"]").click()
    await page.locator("input[name=\"dataAteOspAbertura\"]").fill("31102025") # Modifique a data conforme necessário

    # Espera começar o download
    async with page.expect_download() as download_info:
        # Faz a ação que inicia o download
        await page.get_by_role("button", name=" Excel Completo").click()
    download = await download_info.value
    await download.save_as(downloads_path / download.suggested_filename)
    print("Download de OSP's realizado!")

    print("Todos os arquivos foram baixados!! ✅")

    # ---------------------
    await context.close()
    await browser.close()

async def main() -> None:
    async with async_playwright() as playwright:
        await run(playwright)

asyncio.run(main())
