# Maintaining the README

The root README introduces this project and its setup. Edit `README.md.src`;
`README.md` is committed output. Sebastian-Theme supplies the company badge in
the existing badge row and a compact footer. Project content stays in this repo.

## Set up the contributor tool

Install [mise](https://mise.jdx.dev/getting-started.html) and Git. From a trusted
checkout, run:

```sh
mise trust
mise install --locked
mise run readme:write
mise run readme:check
```

The CLI version belongs to this project in `mise.toml`; `mise.lock` records
release checksums for Linux, macOS, and Windows. README tasks require the
installed pin and never install a missing tool or fall back to a system binary.
This is contributor tooling; people using the project do not need mdtheme.

Edit prose and project badges in `README.md.src`. Keep one ordered pair of
`mdtheme:badges` markers. Do not copy the Sebastian badge or footer into the
source. Review and commit both source and generated output. CI runs the same
read-only check on every pull request and push to the default branch.

## Before pushing

```sh
mise run readme:pre-push
git push
```

The first command regenerates the README and fails if the worktree is dirty,
including untracked files. Review changes, stage and commit them yourself, then
run the command again. It never stages, commits, or pushes. Existing repository
hooks remain in place; this explicit command does not install a hook.

## Update the tool or theme

Update the CLI version in `mise.toml`, then run:

```sh
mise lock --platform linux-arm64,linux-x64,macos-arm64,macos-x64,windows-x64
mise install --locked
mise run readme:write
mise run readme:check
```

Review the lockfile and output together. The theme is pinned separately in
`mdtheme.yaml`. Change its `ref` to a reviewed commit to update branding.
Branches and tags also work; `main` follows branding changes on every run.
Generation and checking need Git access to the theme repository even when the
CLI is already installed. Theme files are data; mdtheme does not execute them.

The [README ownership decision](adr/0001-compose-readme-with-mdtheme.md) records this contract.
