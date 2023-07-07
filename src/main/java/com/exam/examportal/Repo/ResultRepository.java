package com.exam.examportal.Repo;

import com.exam.examportal.model.User;
import com.exam.examportal.model.exam.Quiz;
import com.exam.examportal.model.exam.Result;
import org.springframework.data.jpa.repository.JpaRepository;

import java.util.List;

public interface ResultRepository extends JpaRepository<Result,Integer> {

    List<Result> findByQuiz(Quiz quiz);

    List<Result> findByUser(User user);

    List<Result> findByQuizAndUser(Quiz quiz, User user);
}
