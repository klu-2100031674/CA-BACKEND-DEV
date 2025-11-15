module.exports = {
  testEnvironment: 'node',
  testPathIgnorePatterns: [
    '/node_modules/',
    '/src/python-engine/',
  ],
  modulePathIgnorePatterns: [
    '<rootDir>/src/python-engine/.venv'
  ],
  transform: {},
};
