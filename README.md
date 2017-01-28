# Secret Santa Generator

## Set Up

Requires an xlsx file with columns (in order) of the format:
* name
* email
* helpful hint
* people you don't wish to gift to

## Run

Requires docker.
Build the docker image and run it.

```shell
docker build -t secret-santa . 
docker run secret-santa santa.py $YOURFILE
```

## TODO

Pull out email into own template.
Perhaps allow to customize bot  email? (so don't have to maintain a bot email)


