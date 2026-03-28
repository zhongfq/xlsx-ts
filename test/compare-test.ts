import { runCompareBenchmark } from "../scripts/benchmark.js";

console.log(JSON.stringify(await runCompareBenchmark(), null, 2));
