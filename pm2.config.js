module.exports = {
    apps: [
        {
            name: 'FORMBUILDERAPI',
            script: 'dist/server.js',
            env: {
                PORT: 80,  // Specify the port here
                NODE_ENV: 'production',
            },
            // cwd: '/Server',
        },
    ],
};