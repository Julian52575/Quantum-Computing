{
  description = "Quantum-Computing dev environment: run order_selection_qaoa.ipynb";

  inputs = {
    # Pinned to the 26.05 release tag: https://github.com/NixOS/nixpkgs/releases/tag/26.05
    nixpkgs.url = "github:NixOS/nixpkgs/8c50a710ddca43d7a530fb805ad55bde8d0141c5";
  };

  outputs = { self, nixpkgs }:
    let
      systems = [ "x86_64-linux" "aarch64-linux" "x86_64-darwin" "aarch64-darwin" ];
      forAllSystems = f: nixpkgs.lib.genAttrs systems (system: f system);
    in
    {
      devShells = forAllSystems (system:
        let
          pkgs = nixpkgs.legacyPackages.${system};
          python = pkgs.python312;
        in
        {
          default = pkgs.mkShell {
            packages = [ pkgs.uv python ];

            # Make uv use the Nix interpreter instead of downloading its own.
            UV_PYTHON = "${python}/bin/python";
            UV_PYTHON_DOWNLOADS = "never";

            # PyPI wheels (numpy, qiskit-aer, ...) expect these on a normal FHS system.
            LD_LIBRARY_PATH = pkgs.lib.optionalString pkgs.stdenv.isLinux
              (pkgs.lib.makeLibraryPath [ pkgs.stdenv.cc.cc.lib pkgs.zlib ]);

            shellHook = ''
              uv sync --frozen --quiet
              source .venv/bin/activate
              # Readable notebook diffs/merges in git (repo-local config only).
              nbdime config-git --enable >/dev/null 2>&1
              echo "Notebook env ready. Run: jupyter lab order_selection_qaoa.ipynb"
            '';
          };
        });
    };
}
