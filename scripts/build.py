from pathlib import Path

root = Path(__file__).parent.parent
dist = root / "dist"
dist.mkdir(exist_ok=True)


def load_toml(file: Path) -> dict:
    result = {}
    with file.open() as f:
        for line in f:
            if not line.strip() or line.strip().startswith("#"):
                continue
            if "=" in line:
                key, value = line.split("=", 1)
                key = key.strip()
                value = value.strip().strip('"').strip("'")
                result[key] = value
    return result


def generate_manifest(info: dict) -> str:
    tablen = max(map(len, info.keys()))
    manifest = ["// ==WindhawkMod=="]
    for key, value in info.items():
        manifest.append(f"// @{key:{tablen}} {value}")
    manifest.append("// ==/WindhawkMod==")
    manifest.append("")
    return "\n".join(manifest)


def read_file(path: Path) -> str:
    with path.open() as f:
        return f.read()


def generate_readme() -> str:
    readme = read_file(root / "README.md").split("\n")
    introduction = []
    for line in readme:
        if line.startswith("##"):
            break
        introduction.append(line)
    endl = "\n"
    return f"""// ==WindhawkModDescription==
/*
{endl.join(introduction).strip()}
*/
// ==/WindhawkModReadme=="""


def generate_header() -> str:
    return read_file(root / "scripts" / "header.cpp")


def generate_source() -> str:
    return read_file(root / "main.cpp")


def generate_footer() -> str:
    return read_file(root / "scripts" / "footer.cpp")


def main():
    info = load_toml(root / "project.toml")
    target = [
        generate_manifest(info),
        generate_readme(),
        generate_header(),
        generate_source(),
        generate_footer(),
    ]

    (dist / f"{info['id']}.wh.cpp").write_text("\n".join(target))


if __name__ == "__main__":
    main()
