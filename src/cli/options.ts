export interface CliOptions {
  sprint?: string;
}

export function parseCliOptions(argv: string[]): CliOptions {
  const options: CliOptions = {};
  for (let i = 2; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === '--sprint' || arg === '-s') {
      if (i + 1 >= argv.length) {
        throw new Error('Expected a value after --sprint.');
      }
      options.sprint = argv[++i];
      continue;
    }
    if (arg.startsWith('--sprint=')) {
      options.sprint = arg.slice('--sprint='.length);
      continue;
    }
    if (arg === '--help' || arg === '-h') {
      console.log('Usage: npm run dev -- --sprint "Sprint Name"');
      process.exit(0);
    }
    console.warn(`Ignoring unknown argument: ${arg}`);
  }
  return options;
}
