#!/usr/bin/env python3
"""
초과근무 처리 도구 - Rich 기반 대화형 CLI
"""
import os
import sys
from typing import Dict, Any

from rich.console import Console
from rich.prompt import Prompt, Confirm
from rich.panel import Panel
from rich.progress import Progress, SpinnerColumn, TextColumn, BarColumn
from rich.table import Table
from rich.traceback import install

from utils.config_utils import (
    load_config,
    save_config,
    DOWNLOAD_DIR_KEY,
    WORK_DIR_KEY,
    MEAL_FEE_KEY,
    OFFICIAL_DATA_NAMES_STR_KEY,
)
from main import main as cli_main

# Rich traceback 활성화 (더 예쁜 에러 메시지)
install(show_locals=True)

console = Console()


def show_banner() -> None:
    """프로그램 시작 배너 표시"""
    banner = """
[bold blue]╔══════════════════════════════════════════╗
║    초과근무 처리 도구 v0.1.0            ║
║    Overtime Processing Tool              ║
╚══════════════════════════════════════════╝[/bold blue]
"""
    console.print(banner)


def get_user_inputs(config: Dict[str, Any]) -> Dict[str, str]:
    """사용자로부터 설정값 입력 받기"""
    console.print("\n[bold cyan]설정값을 입력해주세요[/bold cyan]")
    console.print("[dim]기본값이 있는 경우 Enter를 눌러 사용할 수 있습니다.[/dim]\n")

    download_dir = Prompt.ask(
        "📁 [yellow]다운로드 폴더[/yellow]",
        default=config.get(DOWNLOAD_DIR_KEY, "")
    ).strip()

    work_dir = Prompt.ask(
        "💾 [yellow]작업결과 저장 폴더[/yellow]",
        default=config.get(WORK_DIR_KEY, "")
    ).strip()

    meal_fee = Prompt.ask(
        "💰 [yellow]매식비 기준 금액[/yellow]",
        default=config.get(MEAL_FEE_KEY, "5500")
    ).strip()

    names_str = Prompt.ask(
        "👤 [yellow]대상자 이름[/yellow] [dim](쉼표로 구분)[/dim]",
        default=config.get(OFFICIAL_DATA_NAMES_STR_KEY, "")
    ).strip()

    return {
        DOWNLOAD_DIR_KEY: download_dir,
        WORK_DIR_KEY: work_dir,
        MEAL_FEE_KEY: meal_fee,
        OFFICIAL_DATA_NAMES_STR_KEY: names_str
    }


def validate_inputs(inputs: Dict[str, str]) -> bool:
    """입력값 검증"""
    if not inputs[DOWNLOAD_DIR_KEY] or not inputs[WORK_DIR_KEY]:
        console.print("\n[bold red]❌ 오류:[/bold red] DOWNLOAD_DIR과 WORK_DIR은 필수 입력값입니다.")
        return False

    if not os.path.exists(inputs[DOWNLOAD_DIR_KEY]):
        console.print(f"\n[bold red]❌ 오류:[/bold red] 다운로드 폴더가 존재하지 않습니다: {inputs[DOWNLOAD_DIR_KEY]}")
        return False

    # work_dir은 없으면 생성
    if not os.path.exists(inputs[WORK_DIR_KEY]):
        try:
            os.makedirs(inputs[WORK_DIR_KEY], exist_ok=True)
            console.print(f"[green]✓[/green] 작업 폴더 생성: {inputs[WORK_DIR_KEY]}")
        except Exception as e:
            console.print(f"\n[bold red]❌ 오류:[/bold red] 작업 폴더 생성 실패: {e}")
            return False

    return True


def show_config_summary(inputs: Dict[str, str]) -> None:
    """입력된 설정 요약 표시"""
    table = Table(title="설정 요약", show_header=False, box=None)
    table.add_column("항목", style="cyan", width=20)
    table.add_column("값", style="white")

    table.add_row("📁 다운로드 폴더", inputs[DOWNLOAD_DIR_KEY])
    table.add_row("💾 작업 폴더", inputs[WORK_DIR_KEY])
    table.add_row("💰 매식비 기준", inputs[MEAL_FEE_KEY])
    table.add_row("👤 대상자 이름", inputs[OFFICIAL_DATA_NAMES_STR_KEY] or "[dim]없음[/dim]")

    console.print()
    console.print(table)
    console.print()


def run_processing(inputs: Dict[str, str]) -> None:
    """메인 처리 로직 실행"""
    console.print(Panel.fit(
        "[bold green]처리를 시작합니다...[/bold green]",
        border_style="green"
    ))
    console.print()

    try:
        # main 함수 호출
        cli_main(
            download_dir=inputs[DOWNLOAD_DIR_KEY],
            work_dir=inputs[WORK_DIR_KEY],
            meal_fee=inputs[MEAL_FEE_KEY],
            official_data_names_str=inputs[OFFICIAL_DATA_NAMES_STR_KEY]
        )

        console.print()
        console.print(Panel.fit(
            "[bold green]✅ 처리가 완료되었습니다![/bold green]\n\n"
            f"📂 결과 폴더: [cyan]{inputs[WORK_DIR_KEY]}[/cyan]",
            border_style="green",
            title="완료"
        ))

    except Exception as e:
        console.print()
        console.print(Panel.fit(
            f"[bold red]❌ 오류 발생[/bold red]\n\n"
            f"[yellow]{str(e)}[/yellow]",
            border_style="red",
            title="오류"
        ))
        raise


def main() -> None:
    """메인 함수"""
    try:
        # 배너 표시
        show_banner()

        # 설정 로드
        config = load_config()

        # 사용자 입력 받기
        inputs = get_user_inputs(config)

        # 입력값 검증
        if not validate_inputs(inputs):
            sys.exit(1)

        # 설정 요약 표시
        show_config_summary(inputs)

        # 실행 확인
        if not Confirm.ask("[bold]처리를 시작하시겠습니까?[/bold]", default=True):
            console.print("[yellow]취소되었습니다.[/yellow]")
            sys.exit(0)

        # 설정 저장 여부 확인
        save_settings = Confirm.ask(
            "\n[dim]이 설정을 저장하시겠습니까?[/dim]",
            default=True
        )
        if save_settings:
            save_config(inputs)
            console.print("[green]✓[/green] 설정이 저장되었습니다.")

        console.print()

        # 처리 실행
        run_processing(inputs)

    except KeyboardInterrupt:
        console.print("\n\n[yellow]사용자에 의해 중단되었습니다.[/yellow]")
        sys.exit(130)
    except Exception as e:
        console.print(f"\n[bold red]예상치 못한 오류가 발생했습니다:[/bold red] {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
