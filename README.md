# Secret Santa Generator

## Set Up

Requires an xlsx file with columns (in order) of the format:
* name
* email
* people you don't wish to gift to
And then additional columns of survey questions/answers if you wish to give more information to your gifters.  Save this file in the root directory.

## Run

Requires docker.
Build the docker image and run it.

```shell
docker build -t secret-santa . 
```

To send a custom email, create your own Jinja template and put it in the `templates` directory.

Then, run the docker and pass it your xlsx file name, custom (or default) jinja template name, and the email and password of the email account you wish to use to notify participants.

```shell
docker run -v $PWD:/app secret-santa santa.py <xlsx file> <jinja template> <email> <password>
```


