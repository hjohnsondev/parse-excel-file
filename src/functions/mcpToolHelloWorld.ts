import { app, InvocationContext, arg } from "@azure/functions";

export async function mcpToolHello(_toolArguments: any, context: InvocationContext): Promise<string> {
    const mcptoolargs = _toolArguments.arguments as {
        name?: string;
    };
    const name = mcptoolargs?.name;

    console.info(`Hello ${name}, I am MCP Tool!`);

    return `Hello ${name}, I am MCP Tool!`;
}

app.mcpTool('hello', {
    toolName: 'hello',
    description: 'Simple hello world MCP Tool that responses with a hello message.',
    toolProperties: {
      name: arg.string().describe('Name to greet'),
    },
    handler: mcpToolHello
});