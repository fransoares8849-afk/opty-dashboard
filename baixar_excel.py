"""
baixar_excel.py
Faz login no iris.opty.com.br via Autenticação Integrada e exporta o Excel.
Uso: python baixar_excel.py <output_path>
"""

import os
import sys
import time
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

URL_LOGIN  = "https://iris.opty.com.br"
URL_MODULO = "https://iris.opty.com.br/optyparavc/atendimento_exames"
OUTPUT     = sys.argv[1] if len(sys.argv) > 1 else "relatorio-agendamentos.xlsx"

IRIS_EMAIL = os.environ["IRIS_EMAIL"]
IRIS_PASS  = os.environ["IRIS_PASS"]


def run():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        ctx     = browser.new_context(accept_downloads=True)
        page    = ctx.new_page()

        # ── 1. Login ────────────────────────────────────────────────────────
        print("🔐 Abrindo página de login...")
        page.goto(URL_LOGIN, wait_until="networkidle", timeout=30_000)

        # Preenche email + senha (login padrão)
        page.fill('input[type="email"], input[name*="email"], input[placeholder*="mail"]', IRIS_EMAIL)
        page.fill('input[type="password"]', IRIS_PASS)
        page.click('button:has-text("Login"), input[type="submit"]')

        # ── 2. Selecionar módulo ─────────────────────────────────────────────
        print("📦 Selecionando módulo...")
        page.wait_for_selector("text=Selecione o módulo", timeout=15_000)
        page.select_option("select", label="Agendamento de exames")
        # fallback: procura pelo texto se não for <select>
        try:
            page.click('button:has-text("Acessar")')
        except Exception:
            page.get_by_role("button", name="Acessar").click()

        # ── 3. Aguardar carregar a lista ──────────────────────────────────────
        print("⏳ Aguardando carregamento da lista...")
        page.wait_for_url("**/atendimento_exames**", timeout=20_000)
        page.wait_for_load_state("networkidle", timeout=30_000)

        # ── 4. Selecionar todos os status ─────────────────────────────────────
        print("☑️  Selecionando todos os status...")
        try:
            # Tenta abrir o dropdown de status e marcar todos
            status_dropdown = page.locator("text=Status").first
            status_dropdown.click()
            time.sleep(0.5)
            # Tenta clicar em "Selecionar todos" ou marcar cada opção
            try:
                page.click("text=Selecionar todos")
            except Exception:
                # Marca cada status individualmente
                for opt in ["Sem atendimento", "Em atendimento", "Cancelado",
                            "Pendente Doc.", "Pendente Agenda", "Pendente Contato",
                            "Fin. Agendado", "Fin. Pend. Doc.", "Fin. Sem Contato"]:
                    try:
                        page.check(f'input[value*="{opt}"]')
                    except Exception:
                        pass
            # Fecha o dropdown clicando fora
            page.keyboard.press("Escape")
        except Exception as e:
            print(f"   ⚠️  Status dropdown não encontrado, exportando com filtro atual: {e}")

        time.sleep(1)

        # ── 5. Exportar Excel ─────────────────────────────────────────────────
        print("📥 Clicando em Exportar Excel...")
        with page.expect_download(timeout=60_000) as dl_info:
            page.click('button:has-text("Exportar Excel"), a:has-text("Exportar Excel")')

        download = dl_info.value
        download.save_as(OUTPUT)
        print(f"✅ Excel salvo em: {OUTPUT}")

        browser.close()


if __name__ == "__main__":
    run()
