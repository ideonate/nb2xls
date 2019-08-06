- To release a new version of nb2xls on PyPI:

Update __meta__.py (set release version, remove 'dev')

Add version to changelog in README.md

git add the __meta__.py file and git commit

`git tag -a X.X.X -m 'comment'`

`git push`

`git push --tags`

This should be built and pushed to PyPI automatically through travis-ci
