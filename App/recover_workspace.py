"""Local recovery for when the dashboard is unavailable."""
from datetime import datetime
from pathlib import Path

from backup_service import list_backups, restore_backup
from entry_deletion import last_deletion
from entry_restore import restore_entries


def main():
    root = Path(__file__).resolve().parent.parent
    print('Workspace recovery')
    print('Close Excel and the certificate app before restoring files.')
    try:
        latest = last_deletion(root)
        if latest:
            label = (latest.get('supplier') or '1 entry') if latest['count'] == 1 else f"{latest['count']} entries"
            deadline = datetime.fromisoformat(latest['restore_until']).astimezone().strftime('%d %b %Y, %I:%M %p')
            print(f'U. Undo last deletion ({label}, restore by {deadline})')
        try:
            backups = [record for record in list_backups(root) if record['valid']]
        except (OSError, ValueError) as exc:
            backups = []
            print(f'Could not read verified backups: {exc}')
        for index, backup in enumerate(backups, 1):
            print(f"{index}. Restore {backup['created_at']} ({backup['purpose']})")
            print(f"   {backup['path']}")
        if not backups:
            print('No verified workspace backup is available.')
        if not latest and not backups:
            return 1
        choice = input('Choose a recovery action, or press Enter to cancel: ').strip().lower()
        if not choice:
            return 0
        if choice == 'u' and latest:
            selected = [{key:latest[key] for key in ('kind', 'request_id')}]
            preview = restore_entries(root, {'selected':selected})
            for entry in preview['review']['entries']:
                print(entry.get('supplier') or 'Unnamed company')
                if entry.get('restore_sn') and entry['restore_sn'] != entry.get('sn'):
                    print(f"   Restores as SN {entry['restore_sn']}")
                if entry.get('restore_standalone'):
                    print('   Restores as a standalone record.')
            if input('Type UNDO to restore the deleted entries: ').strip() != 'UNDO':
                return 0
            if last_deletion(root) != latest:
                raise ValueError('The latest deletion changed. Reopen recovery and check the new details.')
            restore_entries(root, {'selected':selected, 'review_token':preview['review_token'], 'checked':True}, commit=True)
            print('Deletion restored. Reopen Prepare Certificates to check the company list.')
        else:
            if not choice.isascii() or not choice.isdecimal() or not 1 <= int(choice) <= len(backups):
                raise ValueError('Choose an action from the list.')
            backup = backups[int(choice) - 1]
            print(f"Selected backup: {backup['path']}")
            print('This restores workbooks and certificate data to the selected backup. A copy of the current state is saved first.')
            if input('Type RESTORE to confirm: ').strip() != 'RESTORE':
                return 0
            result = restore_backup(root, backup['backup_id'], confirmed=True)
            print(f"Workspace restored. The previous state was saved to {result['pre_restore_path']}")
            for warning in result.get('warnings', ()):
                print(warning)
        input('Press Enter to close.')
        return 0
    except (EOFError, KeyboardInterrupt):
        return 0
    except (OSError, ValueError, KeyError) as exc:
        print(f'Recovery could not finish: {exc}')
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
