#!/usr/bin/env node
/**
 * m365-mcp auth CLI — manage accounts and authenticate outside the MCP server
 *
 * Usage:
 *   npx tsx src/auth-cli.ts add     <name> <tenantId> <clientId> [email] [desc]
 *   npx tsx src/auth-cli.ts remove  <name>
 *   npx tsx src/auth-cli.ts default <name>
 *   npx tsx src/auth-cli.ts list
 *   npx tsx src/auth-cli.ts login   [--account=name]
 *   npx tsx src/auth-cli.ts status  [--account=name]
 */

import {
  addAccount,
  removeAccount,
  setDefaultAccount,
  listAccounts,
  authenticate,
  getTokenStatus,
} from "./auth.js";

async function main() {
  const args = process.argv.slice(2);
  const cmd = args[0];
  const rest = args.slice(1);
  const accountArg = rest.find((a) => a.startsWith("--account="));
  const accountName = accountArg ? accountArg.split("=")[1] : undefined;
  const positional = rest.filter((a) => !a.startsWith("--"));

  try {
    switch (cmd) {
      case "add": {
        if (positional.length < 3) {
          console.log("Usage: m365-mcp auth add <name> <tenantId> <clientId> [email] [description]");
          process.exit(1);
        }
        const [name, tenantId, clientId, email, desc] = positional;
        addAccount(name, tenantId, clientId, email, desc);
        console.log(`✅ Account "${name}" added.`);
        break;
      }

      case "remove": {
        if (positional.length < 1) {
          console.log("Usage: m365-mcp auth remove <name>");
          process.exit(1);
        }
        removeAccount(positional[0]);
        console.log(`✅ Account "${positional[0]}" removed.`);
        break;
      }

      case "default": {
        if (positional.length < 1) {
          console.log("Usage: m365-mcp auth default <name>");
          process.exit(1);
        }
        setDefaultAccount(positional[0]);
        console.log(`✅ Default account set to "${positional[0]}".`);
        break;
      }

      case "list": {
        const { default: def, accounts } = listAccounts();
        if (accounts.length === 0) {
          console.log("No accounts configured.\n");
          console.log("Add one: m365-mcp auth add <name> <tenantId> <clientId>");
          break;
        }
        console.log("📧 M365 Accounts:\n");
        for (const acc of accounts) {
          const marker = acc.isDefault ? " [DEFAULT]" : "";
          console.log(`  ${acc.name}${marker}`);
          if (acc.email) console.log(`    Email: ${acc.email}`);
          if (acc.description) console.log(`    Desc:  ${acc.description}`);
          console.log(`    Tenant: ${acc.tenantId}`);
          console.log();
        }
        break;
      }

      case "login": {
        await authenticate(accountName, (url, code) => {
          console.log("\n" + "=".repeat(60));
          console.log("🔐 Sign in to Microsoft");
          console.log("=".repeat(60));
          console.log(`\n1. Open: ${url}`);
          console.log(`2. Enter code: ${code}`);
          console.log(`3. Sign in and approve permissions\n`);
        });
        console.log("✅ Authenticated successfully.");
        break;
      }

      case "status": {
        const status = await getTokenStatus(accountName);
        if (status.valid) {
          console.log(`✅ ${status.account}: Authenticated (expires ${status.expires})`);
        } else {
          console.log(`❌ ${status.account}: Not authenticated`);
          process.exit(1);
        }
        break;
      }

      case "env": {
        // Dump env vars needed for OpenClaw setup instructions
        const { accounts } = listAccounts();
        if (accounts.length === 0) {
          console.log("# No accounts configured.");
          break;
        }
        const acc = accounts.find((a) => a.isDefault) || accounts[0];
        console.log(`M365_ACCOUNT=${acc.name}`);
        console.log(`# M365_TIMEZONE=UTC  (optional, default: UTC)`);
        console.log(`# Run 'm365-mcp auth login --account=${acc.name}' to authenticate`);
        break;
      }

      default:
        console.log("m365-mcp auth — Account Management CLI\n");
        console.log("Commands:");
        console.log("  add     <name> <tenantId> <clientId> [email] [desc]");
        console.log("  remove  <name>");
        console.log("  default <name>");
        console.log("  list");
        console.log("  login   [--account=name]");
        console.log("  status  [--account=name]");
        console.log("  env");
        console.log("\nSetup guide: https://github.com/sam2kb/m365-mcp#readme");
        process.exit(1);
    }
  } catch (err: any) {
    console.error(`❌ ${err.message}`);
    process.exit(1);
  }
}

main();
