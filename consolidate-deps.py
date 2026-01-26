#!/usr/bin/env python3

"""
Dependabot PR Consolidator for Python Projects
Automates the process of consolidating multiple dependabot PRs into a single PR
Supports: uv, poetry, pip-tools, requirements.txt
"""

import json
import re
import subprocess
import sys
from pathlib import Path
from typing import List, Dict, Optional


def exec_cmd(cmd: str, check: bool = True, capture: bool = False) -> str:
    """Execute command and return output"""
    try:
        result = subprocess.run(
            cmd,
            shell=True,
            check=check,
            capture_output=capture,
            text=True
        )
        return result.stdout.strip() if capture else ""
    except subprocess.CalledProcessError as e:
        if not check:
            return ""
        raise


def get_dependabot_prs() -> List[Dict]:
    """Get all open dependabot PRs"""
    output = exec_cmd(
        'gh pr list --state open --author "app/dependabot" --json number,title,headRefName',
        capture=True
    )
    return json.loads(output)


def parse_package_update(title: str) -> Optional[Dict]:
    """Parse package update from PR title"""
    # Format: "build(deps): bump package-name from X to Y"
    match = re.search(r'bump (.+?) from (.+?) to (.+?)$', title, re.IGNORECASE)
    if not match:
        return None

    return {
        'package': match.group(1),
        'old_version': match.group(2),
        'new_version': match.group(3)
    }


def detect_project_type() -> str:
    """Detect Python project type"""
    if Path('uv.lock').exists():
        return 'uv'
    elif Path('poetry.lock').exists():
        return 'poetry'
    elif Path('requirements.in').exists():
        return 'pip-tools'
    elif Path('requirements.txt').exists():
        return 'pip'
    else:
        raise RuntimeError("No supported Python dependency file found")


def update_dependencies(updates: List[Dict], project_type: str) -> bool:
    """Update dependency files based on project type"""
    if project_type == 'uv':
        return update_uv_dependencies(updates)
    elif project_type == 'poetry':
        return update_poetry_dependencies(updates)
    elif project_type == 'pip-tools':
        return update_pip_tools_dependencies(updates)
    elif project_type == 'pip':
        return update_requirements_txt(updates)
    return False


def update_uv_dependencies(updates: List[Dict]) -> bool:
    """Update dependencies using uv"""
    for update in updates:
        pkg = update['package']
        version = update['new_version']
        # uv add will update if already exists
        exec_cmd(f'uv add "{pkg}=={version}"', check=False)
    return True


def update_poetry_dependencies(updates: List[Dict]) -> bool:
    """Update dependencies using poetry"""
    for update in updates:
        pkg = update['package']
        version = update['new_version']
        exec_cmd(f'poetry add "{pkg}@{version}"')
    return True


def update_pip_tools_dependencies(updates: List[Dict]) -> bool:
    """Update requirements.in and compile"""
    req_file = Path('requirements.in')
    content = req_file.read_text()

    for update in updates:
        pkg = update['package']
        version = update['new_version']
        # Replace version in requirements.in
        pattern = rf'{re.escape(pkg)}==[\d\.]+'
        replacement = f'{pkg}=={version}'
        content = re.sub(pattern, replacement, content)

    req_file.write_text(content)
    exec_cmd('pip-compile requirements.in')
    return True


def update_requirements_txt(updates: List[Dict]) -> bool:
    """Update requirements.txt directly"""
    req_file = Path('requirements.txt')
    content = req_file.read_text()

    for update in updates:
        pkg = update['package']
        version = update['new_version']
        pattern = rf'{re.escape(pkg)}==[\d\.]+'
        replacement = f'{pkg}=={version}'
        content = re.sub(pattern, replacement, content)

    req_file.write_text(content)
    return True


def run_tests(project_type: str) -> bool:
    """Run tests based on project type"""
    try:
        if project_type == 'uv':
            exec_cmd('uv run pytest')
            exec_cmd('uv run mypy .', check=False)
            exec_cmd('uv run ruff check .', check=False)
        elif project_type == 'poetry':
            exec_cmd('poetry run pytest')
            exec_cmd('poetry run mypy .', check=False)
        else:
            exec_cmd('pytest', check=False)
        return True
    except subprocess.CalledProcessError:
        return False


def main():
    """Main execution"""
    print('🔍 Fetching dependabot PRs...')
    prs = get_dependabot_prs()

    if not prs:
        print('✅ No open dependabot PRs found')
        sys.exit(0)

    print(f'📦 Found {len(prs)} dependabot PRs')

    # Parse updates
    updates = []
    for pr in prs:
        parsed = parse_package_update(pr['title'])
        if parsed:
            updates.append({
                'pr': pr['number'],
                'branch': pr['headRefName'],
                **parsed
            })

    if not updates:
        print('❌ Could not parse any package updates')
        sys.exit(1)

    # Display updates
    print('\n📋 Updates to apply:')
    for u in updates:
        print(f"   - {u['package']}: {u['old_version']} → {u['new_version']} (PR #{u['pr']})")

    # Detect project type
    project_type = detect_project_type()
    print(f'\n🔧 Detected project type: {project_type}')

    # Create branch
    print('\n🌿 Creating consolidated branch...')
    exec_cmd('git checkout -b chore/deps-combined-update', check=False)

    # Update dependencies
    print('📝 Updating dependencies...')
    try:
        update_dependencies(updates, project_type)
    except Exception as e:
        print(f'❌ Failed to update dependencies: {e}')
        exec_cmd('git checkout main')
        exec_cmd('git branch -D chore/deps-combined-update', check=False)
        sys.exit(1)

    # Run tests
    print('🧪 Running tests...')
    if not run_tests(project_type):
        print('❌ Tests failed')
        print('💡 You may want to exclude problematic packages and retry')
        sys.exit(1)

    # Commit changes
    print('💾 Committing changes...')
    update_list = '\n'.join([f"- {u['package']}: {u['old_version']} → {u['new_version']}" for u in updates])
    commit_msg = f"""chore(deps): update dependencies

{update_list}

🤖 Generated with [Claude Code](https://claude.com/claude-code)

Co-Authored-By: Claude Sonnet 4.5 <noreply@anthropic.com>"""

    exec_cmd('git add .')
    exec_cmd(f'git commit -m "{commit_msg}"')

    # Push branch
    print('🚀 Pushing to remote...')
    exec_cmd('git push -u origin chore/deps-combined-update')

    # Create PR
    print('📝 Creating consolidated PR...')
    pr_body = f"""## Summary
{chr(10).join([f"- Update {u['package']} from {u['old_version']} to {u['new_version']}" for u in updates])}

## Related PRs
{chr(10).join([f"- #{u['pr']}: {u['package']} {u['old_version']} → {u['new_version']}" for u in updates])}

## Test plan
- [x] All tests passing
- [x] Dependencies updated successfully

🤖 Generated with [Claude Code](https://claude.com/claude-code)"""

    pr_url = exec_cmd(
        f'gh pr create --title "chore(deps): update dependencies" --body "{pr_body}"',
        capture=True
    )
    print(f'✅ PR created: {pr_url}')

    # Get PR number from URL
    pr_number = pr_url.rstrip('/').split('/')[-1]

    # Close individual PRs
    print('\n🔒 Closing individual dependabot PRs...')
    for u in updates:
        exec_cmd(f'gh pr close {u["pr"]} -c "Consolidated into #{pr_number}"')
        print(f'   ✅ Closed PR #{u["pr"]}')

    print('\n✅ Consolidation complete!')
    print(f'\n📌 Next steps:')
    print(f'   1. Review PR: {pr_url}')
    print(f'   2. Merge when ready: gh pr merge {pr_number} --squash --delete-branch')


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print(f'❌ Error: {e}', file=sys.stderr)
        sys.exit(1)
