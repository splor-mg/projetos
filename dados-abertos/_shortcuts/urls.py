import argparse
import glob
import json
import os
import win32com.client
import subprocess

def read_url_file(file_path):
    """Read a .url file and return a dictionary with title and url."""
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
        # Extract URL from the file content
        url_start = content.find('URL=') + 4
        url_end = content.find('\n', url_start)
        url = content[url_start:url_end].strip()
        # Use filename as title (without .url extension)
        title = os.path.splitext(os.path.basename(file_path))[0]
        return {'title': title, 'url': url}

def sanitize_filename(filename):
    """Replace invalid filename characters with underscores."""
    # List of invalid characters in Windows filenames
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        filename = filename.replace(char, '_')
    return filename

def create_url_file(bookmark, shortcuts_dir):
    """Create a .url file from a bookmark dictionary."""
    title = sanitize_filename(bookmark['title'])
    url = bookmark['url']
    file_path = os.path.join(shortcuts_dir, f"{title}.url")
    
    content = f"""[InternetShortcut]
URL={url}
"""
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(content)

def read_favorites_html(file_path):
    """Read a favorites.html file and return a list of bookmark dictionaries."""
    bookmarks = []
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
        # Simple parsing of HTML bookmarks
        # This is a basic implementation and might need to be enhanced
        import re
        pattern = r'<A HREF="([^"]+)"[^>]*>([^<]+)</A>'
        matches = re.finditer(pattern, content)
        for match in matches:
            url = match.group(1)
            title = match.group(2)
            bookmarks.append({'title': title, 'url': url})
    return bookmarks

def create_favorites_html(bookmarks, output_path):
    """Create a favorites.html file from a list of bookmark dictionaries."""
    content = """<!DOCTYPE NETSCAPE-Bookmark-file-1>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=UTF-8">
<TITLE>Bookmarks</TITLE>
<H1>Bookmarks</H1>
<DL><p>
    <DT><H3>Bookmarks</H3>
    <DL><p>
"""
    
    for bookmark in bookmarks:
        content += f'        <DT><A HREF="{bookmark["url"]}">{bookmark["title"]}</A>\n'
    
    content += """    </DL><p>
</DL><p>
"""
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(content)

def check_consistency(shortcuts_dir):
    """Check consistency between .url files, urls_list.json, and favorites.html."""
    # Get all .url files
    url_files = glob.glob(os.path.join(shortcuts_dir, '*.url'))
    url_bookmarks = [read_url_file(f) for f in url_files]
    url_titles = {b['title'] for b in url_bookmarks}
    
    has_differences = False
    
    # Check urls_list.json if it exists
    json_path = os.path.join(shortcuts_dir, 'urls_list.json')
    if os.path.exists(json_path):
        with open(json_path, 'r', encoding='utf-8') as f:
            json_bookmarks = json.load(f)
            json_titles = {b['title'] for b in json_bookmarks}
            
            # Compare with .url files
            if url_titles != json_titles:
                has_differences = True
                print("Warning: urls_list.json and .url files are different")
                print("In .url but not in json:", url_titles - json_titles)
                print("In json but not in .url:", json_titles - url_titles)
    
    # Check favorites.html if it exists
    fav_path = os.path.join(shortcuts_dir, 'favorites.html')
    if os.path.exists(fav_path):
        fav_bookmarks = read_favorites_html(fav_path)
        fav_titles = {b['title'] for b in fav_bookmarks}
        
        # Compare with .url files
        if url_titles != fav_titles:
            has_differences = True
            print("Warning: favorites.html and .url files are different")
            print("In .url but not in favorites.html:", url_titles - fav_titles)
            print("In favorites.html but not in .url:", fav_titles - url_titles)
    
    if not has_differences:
        print("OK: All files are consistent")

def get_current_repo_name():
    """Get the current repository name from git."""
    try:
        result = subprocess.run(['git', 'rev-parse', '--show-toplevel'], 
                              capture_output=True, text=True, check=True)
        return os.path.basename(result.stdout.strip())
    except subprocess.CalledProcessError:
        return None

def create_shortcut(target_path, shortcut_path):
    """Create a Windows shortcut (.lnk) file."""
    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortCut(shortcut_path)
    shortcut.Targetpath = target_path
    shortcut.WorkingDirectory = target_path
    shortcut.save()

def create_single_url(short_name, shortcuts_dir):
    """Create a single .url file with template content."""
    file_path = os.path.join(shortcuts_dir, f"{short_name}.url")
    content = f"""[InternetShortcut]
URL=https://www.{short_name}.com.br/
"""
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(content)
    print(f"Created {file_path}")

def create_template_json(shortcuts_dir):
    """Create a template urls_list.json file."""
    file_path = os.path.join(shortcuts_dir, 'urls_list.json')
    
    # Get current repository name
    repo_name = get_current_repo_name()
    flow_panel_url = f"https://github.com/users/carloshob/projects/1/views/3?sliceBy%5Bvalue%5D=carloshob%2F{repo_name}" if repo_name else "https://github.com/users/carloshob/projects/1/views/3"
    
    template = [
        {
            "title": "flow-panel",
            "url": flow_panel_url
        },
        {
            "title": "gst-vista",
            "url": "https://github.com/orgs/splor-mg/projects/13/views/17"
        }
    ]
    with open(file_path, 'w', encoding='utf-8') as f:
        json.dump(template, f, indent=2)
    print(f"Created {file_path}")

def create_workspace_shortcut(workspace_type):
    """Create a shortcut to the current repository in specified workspace directory."""
    repo_name = get_current_repo_name()
    if not repo_name:
        print("Error: Not in a git repository")
        return

    current_dir = os.getcwd()
    workspace_dirs = {
        'atd': r"C:\Users\Carlos\OneDrive\!pss\!atd",
        'atdt': r"C:\Users\Carlos\OneDrive\!trb\!atd-t",
        'dev': r"C:\Users\Carlos\OneDrive\!pss\!dev",
        'est': r"C:\Users\Carlos\OneDrive\!pss\!est"
    }
    
    target_dir = workspace_dirs.get(workspace_type)
    if not target_dir:
        print(f"Error: Unknown workspace type: {workspace_type}")
        return

    shortcut_path = os.path.join(target_dir, f"{repo_name}.lnk")
    create_shortcut(current_dir, shortcut_path)
    print(f"Created shortcut at {shortcut_path}")

def create_flow_panel_url(shortcuts_dir):
    """Create a .url file for the GitHub flow panel of the current repository."""
    repo_name = get_current_repo_name()
    if not repo_name:
        print("Error: Not in a git repository")
        return

    file_path = os.path.join(shortcuts_dir, "flow-panel.url")
    url = f"https://github.com/users/carloshob/projects/1/views/3?sliceBy%5Bvalue%5D=carloshob%2F{repo_name}"
    
    content = f"""[InternetShortcut]
URL={url}
"""
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(content)
    print(f"Created {file_path}")

def create_aid_url(shortcuts_dir):
    """Create a .url file for the gst-vista project."""
    file_path = os.path.join(shortcuts_dir, "gst-vista.url")
    url = "https://github.com/orgs/splor-mg/projects/13/views/17"
    
    content = f"""[InternetShortcut]
URL={url}
"""
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(content)
    print(f"Created {file_path}")

def main():
    parser = argparse.ArgumentParser(
        description='URL Shortcuts Manager - A tool to manage and convert between different URL shortcut formats',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Create a single .url file with template content
  _shortcuts/urls.py create --short my-shortcut

  # Create a template urls_list.json file
  _shortcuts/urls.py create --list

  # Create workspace shortcuts for current repository
  _shortcuts/urls.py create --atd     # Create ATD shortcut
  _shortcuts/urls.py create --atdt    # Create ATD-T shortcut
  _shortcuts/urls.py create --dev     # Create DEV shortcut
  _shortcuts/urls.py create --est     # Create EST shortcut
  _shortcuts/urls.py create --aid     # Create gst-vista shortcut

  # Create GitHub flow panel shortcut for current repository
  _shortcuts/urls.py create --flow

  # Create urls_list.json from .url files
  _shortcuts/urls.py create --short2list

  # Create .url files from urls_list.json
  _shortcuts/urls.py create --list2short

  # Create favorites.html from .url files
  _shortcuts/urls.py create --short2fav

  # Check consistency between all files
  _shortcuts/urls.py check
"""
    )
    
    subparsers = parser.add_subparsers(dest='command', help='Available commands')
    
    # Create command
    create_parser = subparsers.add_parser('create', help='Create or convert between different URL shortcut formats')
    create_group = create_parser.add_argument_group('Conversion Options')
    
    # New single shortcut options
    create_group.add_argument('--short', metavar='SHORT_NAME',
        help='Create a single .url file with template content')
    create_group.add_argument('--list', action='store_true',
        help='Create a template urls_list.json file')
    create_group.add_argument('--flow', action='store_true',
        help='Create GitHub flow panel shortcut for current repository')
    create_group.add_argument('--aid', action='store_true',
        help='Create gst-vista shortcut')
    create_group.add_argument('--atd', action='store_true',
        help='Create ATD shortcut for current repository')
    create_group.add_argument('--atdt', action='store_true',
        help='Create ATD-T shortcut for current repository')
    create_group.add_argument('--dev', action='store_true',
        help='Create DEV shortcut for current repository')
    create_group.add_argument('--est', action='store_true',
        help='Create EST shortcut for current repository')
    
    # Existing conversion options
    create_group.add_argument('--short2list', action='store_true', 
        help='Create urls_list.json from .url files')
    create_group.add_argument('--fav2list', action='store_true', 
        help='Create urls_list.json from favorites.html')
    create_group.add_argument('--list2short', action='store_true', 
        help='Create .url files from urls_list.json')
    create_group.add_argument('--fav2short', action='store_true', 
        help='Create .url files from favorites.html')
    create_group.add_argument('--short2fav', action='store_true', 
        help='Create favorites.html from .url files')
    create_group.add_argument('--list2fav', action='store_true', 
        help='Create favorites.html from urls_list.json')
    
    # Check command
    check_parser = subparsers.add_parser('check', help='Check consistency between different URL shortcut formats')
    
    args = parser.parse_args()
    
    # Get the directory where the script is located
    shortcuts_dir = os.path.dirname(os.path.abspath(__file__))
    
    if args.command == 'create':
        if not any([args.short, args.list, args.flow, args.aid, args.atd, args.atdt, args.dev, args.est,
                   args.short2list, args.fav2list, args.list2short, 
                   args.fav2short, args.short2fav, args.list2fav]):
            create_parser.print_help()
            return
            
        if args.short:
            create_single_url(args.short, shortcuts_dir)
        elif args.list:
            create_template_json(shortcuts_dir)
        elif args.flow:
            create_flow_panel_url(shortcuts_dir)
        elif args.aid:
            create_aid_url(shortcuts_dir)
        elif args.atd:
            create_workspace_shortcut('atd')
        elif args.atdt:
            create_workspace_shortcut('atdt')
        elif args.dev:
            create_workspace_shortcut('dev')
        elif args.est:
            create_workspace_shortcut('est')
        elif args.short2list:
            url_files = glob.glob(os.path.join(shortcuts_dir, '*.url'))
            bookmarks = [read_url_file(f) for f in url_files]
            with open(os.path.join(shortcuts_dir, 'urls_list.json'), 'w', encoding='utf-8') as f:
                json.dump(bookmarks, f, indent=2)
            print("Created urls_list.json from .url files")
            
        elif args.fav2list:
            fav_path = os.path.join(shortcuts_dir, 'favorites.html')
            if os.path.exists(fav_path):
                bookmarks = read_favorites_html(fav_path)
                with open(os.path.join(shortcuts_dir, 'urls_list.json'), 'w', encoding='utf-8') as f:
                    json.dump(bookmarks, f, indent=2)
                print("Created urls_list.json from favorites.html")
            else:
                print("Error: favorites.html not found")
                
        elif args.list2short:
            json_path = os.path.join(shortcuts_dir, 'urls_list.json')
            if os.path.exists(json_path):
                with open(json_path, 'r', encoding='utf-8') as f:
                    bookmarks = json.load(f)
                for bookmark in bookmarks:
                    create_url_file(bookmark, shortcuts_dir)
                print("Created .url files from urls_list.json")
            else:
                print("Error: urls_list.json not found")
                
        elif args.fav2short:
            fav_path = os.path.join(shortcuts_dir, 'favorites.html')
            if os.path.exists(fav_path):
                bookmarks = read_favorites_html(fav_path)
                for bookmark in bookmarks:
                    create_url_file(bookmark, shortcuts_dir)
                print("Created .url files from favorites.html")
            else:
                print("Error: favorites.html not found")
                
        elif args.short2fav:
            url_files = glob.glob(os.path.join(shortcuts_dir, '*.url'))
            bookmarks = [read_url_file(f) for f in url_files]
            create_favorites_html(bookmarks, os.path.join(shortcuts_dir, 'favorites.html'))
            print("Created favorites.html from .url files")
            
        elif args.list2fav:
            json_path = os.path.join(shortcuts_dir, 'urls_list.json')
            if os.path.exists(json_path):
                with open(json_path, 'r', encoding='utf-8') as f:
                    bookmarks = json.load(f)
                create_favorites_html(bookmarks, os.path.join(shortcuts_dir, 'favorites.html'))
                print("Created favorites.html from urls_list.json")
            else:
                print("Error: urls_list.json not found")
                
    elif args.command == 'check':
        print("Checking consistency between files...")
        check_consistency(shortcuts_dir)
    else:
        parser.print_help()

if __name__ == '__main__':
    main() 