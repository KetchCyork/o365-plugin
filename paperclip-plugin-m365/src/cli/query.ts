/** query — run the m365 executor once from the CLI. Usage: npm run query -- mail "budget" */
import "dotenv/config";
import executor from "../executor.js";

async function main() {
  const [intent = "briefing", ...rest] = process.argv.slice(2);
  const query = rest.join(" ");
  const res = await executor.run(
    { title: intent, body: query, raw: { intent, query } },
    {}
  );
  console.log(res.output);
}
main().catch((e) => { console.error(e); process.exit(1); });
