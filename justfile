install:
	npm install

clean:
	rm -rf dist

test:
	npx vitest --run

build: install clean
	npx tsc -p tsconfig.lib.json
	rm dist/tsconfig.lib.tsbuildinfo

publish: test build
	npm publish --access public