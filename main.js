const { spawn, spawnSync } = require("child_process");

function run(command, args, opts = {}) {
  return new Promise((resolve, reject) => {
    const proc = spawn(command, args, {
      stdio: "inherit",
      shell: true,
      ...opts,
    });
    proc.on("exit", (code) => {
      if (code === 0) resolve();
      else
        reject(
          new Error(`${command} ${args.join(" ")} failed with code ${code}`)
        );
    });
  });
}

async function main() {
  await run("npm", ["run", "dev-cert"]);
  const server = spawn("npm", ["start"], { stdio: "inherit", shell: true });
  await new Promise((r) => setTimeout(r, 1500));
  await run("npm", ["run", "sideload"]);
  console.log("Add-in sideloaded. Press Ctrl+C to stop.");
  const onExit = () => {
    spawnSync("npm", ["run", "unsideload"], { stdio: "inherit", shell: true });
    try {
      server.kill();
    } catch {}
    process.exit(0);
  };
  process.on("SIGINT", onExit);
  process.on("SIGTERM", onExit);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
