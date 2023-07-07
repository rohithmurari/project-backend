package com.exam.examportal.service.Impl;

import com.exam.examportal.Repo.ResultRepository;
import com.exam.examportal.model.User;
import com.exam.examportal.model.exam.Quiz;
import com.exam.examportal.model.exam.Result;
import com.exam.examportal.service.ResultService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.List;
@Service
public class ResultServiceImpl implements ResultService {

    @Autowired
    private ResultRepository resultrepository;

    @Override
    public Result addResult(Result result) {
        return resultrepository.save(result);
    }

    @Override
    public List<Result> getAllResult() {
        return resultrepository.findAll();
    }

    @Override
    public List<Result> getResultOfQuiz(Quiz quiz) {
        return resultrepository.findByQuiz(quiz);
    }

    @Override
    public List<Result> getResultOfUser(User user) {
        return resultrepository.findByUser(user);
    }

    @Override
    public List<Result> getResultOfUserAndQuiz(Quiz quiz, User user) {
        return resultrepository.findByQuizAndUser(quiz,user);
    }
}
